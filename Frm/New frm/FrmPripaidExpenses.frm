VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmPripaidExpenses 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19260
   Icon            =   "FrmPripaidExpenses.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   19260
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   3795
      Left            =   0
      TabIndex        =   53
      Top             =   3600
      Width           =   19155
      _cx             =   33787
      _cy             =   6694
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
      BackColorAlternate=   16777088
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
      Cols            =   33
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmPripaidExpenses.frx":6852
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   1200
         TabIndex        =   54
         Top             =   960
         Visible         =   0   'False
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   21000
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmPripaidExpenses.frx":6D62
      Left            =   20880
      List            =   "FrmPripaidExpenses.frx":6D72
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   39
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
      TabIndex        =   33
      Top             =   0
      Width           =   19305
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   34
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
         ButtonImage     =   "FrmPripaidExpenses.frx":6D8B
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   35
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
         ButtonImage     =   "FrmPripaidExpenses.frx":7125
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   36
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
         ButtonImage     =   "FrmPripaidExpenses.frx":74BF
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   37
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
         ButtonImage     =   "FrmPripaidExpenses.frx":7859
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄—Ì› «·„’—Ê›«  «·„ﬁœ„…"
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
         TabIndex        =   38
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmPripaidExpenses.frx":7BF3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4215
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   720
      Width           =   19275
      Begin VB.CheckBox ChkAll 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   17760
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   2055
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   19335
         Begin VB.TextBox TxtAccount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   16950
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   1320
            Width           =   705
         End
         Begin VB.TextBox TxtRemark2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1320
            Width           =   5415
         End
         Begin VB.TextBox TxtValu 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox TxtCustCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   16950
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   600
            Width           =   705
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„’—Ê›"
            Height          =   1215
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   120
            Width           =   2535
            Begin XtremeSuiteControls.CheckBox ChMessier 
               Height          =   255
               Left            =   360
               TabIndex        =   79
               Top             =   840
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "ÌœŒ· ›Ì «·„”Ì—"
               ForeColor       =   16711680
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Opt 
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   5
               Top             =   600
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Õ”«»"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Opt 
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   6
               Top             =   600
               Width           =   975
               _Version        =   786432
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„ÊŸ›"
               ForeColor       =   0
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.TextBox TxtNameE 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox TxtName 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   16950
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   960
            Width           =   705
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Bindings        =   "FrmPripaidExpenses.frx":8FF8
            Height          =   315
            Left            =   9240
            TabIndex        =   10
            Top             =   960
            Width           =   7695
            _ExtentX        =   13573
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
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmPripaidExpenses.frx":900D
            Height          =   315
            Left            =   9240
            TabIndex        =   1
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
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
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   1320
            Width           =   2535
            _ExtentX        =   4471
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
            ButtonImage     =   "FrmPripaidExpenses.frx":9022
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   9240
            TabIndex        =   8
            Top             =   600
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   315
            Left            =   6000
            TabIndex        =   13
            Top             =   960
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Format          =   98041857
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker HistoryDate 
            Height          =   315
            Left            =   2760
            TabIndex        =   12
            Top             =   600
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Format          =   98041857
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   315
            Left            =   2760
            TabIndex        =   14
            Top             =   960
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Format          =   98041857
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbAccount 
            Height          =   315
            Left            =   9240
            TabIndex        =   76
            Top             =   1320
            Width           =   7695
            _ExtentX        =   13573
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcCostCenter 
            Bindings        =   "FrmPripaidExpenses.frx":F884
            Height          =   315
            Left            =   2640
            TabIndex        =   81
            Top             =   1680
            Width           =   5550
            _ExtentX        =   9790
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
         Begin MSDataListLib.DataCombo dcproject 
            Height          =   315
            Left            =   9240
            TabIndex        =   83
            Top             =   1680
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘—Ê⁄"
            Height          =   255
            Index           =   15
            Left            =   17760
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„—ﬂ“ «· ﬂ·›…  "
            Height          =   270
            Left            =   8010
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «·„’—Ê›"
            Height          =   285
            Index           =   14
            Left            =   17520
            TabIndex        =   77
            Top             =   1320
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            Height          =   255
            Index           =   13
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   285
            Index           =   11
            Left            =   4800
            TabIndex        =   71
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰  √—ÌŒ"
            Height          =   285
            Index           =   6
            Left            =   8160
            TabIndex        =   70
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«À»« "
            Height          =   285
            Index           =   1
            Left            =   4800
            TabIndex        =   69
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„…"
            Height          =   285
            Index           =   0
            Left            =   8160
            TabIndex        =   68
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÊŸ›"
            Height          =   285
            Index           =   10
            Left            =   17520
            TabIndex        =   67
            Top             =   960
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ‹ „œ›Ê⁄«  „ﬁœ„…"
            Height          =   285
            Index           =   5
            Left            =   17520
            TabIndex        =   66
            Top             =   600
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ ≈‰Ã·Ì“Ì"
            Height          =   285
            Index           =   9
            Left            =   4800
            TabIndex        =   65
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   17520
            TabIndex        =   64
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ ⁄—»Ì"
            Height          =   285
            Index           =   3
            Left            =   8160
            TabIndex        =   60
            Top             =   240
            Width           =   1245
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   5280
         TabIndex        =   27
         Top             =   120
         Width           =   14055
         Begin VB.TextBox TxtRemark 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   240
            Width           =   7575
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   8760
            TabIndex        =   29
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   98041857
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   285
            Index           =   4
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            Height          =   255
            Index           =   12
            Left            =   8040
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   10410
            TabIndex        =   30
            Top             =   255
            Width           =   885
         End
      End
      Begin XtremeSuiteControls.RadioButton Opttype 
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   85
         Top             =   360
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ÃœÌœ"
         ForeColor       =   0
         BackColor       =   14871017
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Opttype 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   86
         Top             =   360
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«›  «ÕÌ"
         ForeColor       =   255
         BackColor       =   14871017
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   21000
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   21000
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   20640
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   21240
      TabIndex        =   41
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
      Left            =   20880
      TabIndex        =   42
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
      Left            =   1920
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   6960
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   45
         Top             =   480
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   44
         Top             =   1080
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12960
            TabIndex        =   18
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   240
            Width           =   990
            _ExtentX        =   1746
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
            ButtonImage     =   "FrmPripaidExpenses.frx":F899
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   10200
            TabIndex        =   20
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            ButtonImage     =   "FrmPripaidExpenses.frx":160FB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11520
            TabIndex        =   19
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
            ButtonImage     =   "FrmPripaidExpenses.frx":16495
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   8880
            TabIndex        =   21
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            ButtonImage     =   "FrmPripaidExpenses.frx":1CCF7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   7560
            TabIndex        =   22
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
            ButtonImage     =   "FrmPripaidExpenses.frx":1D091
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   23
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
            ButtonImage     =   "FrmPripaidExpenses.frx":1D62B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   4320
            TabIndex        =   61
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
            ButtonImage     =   "FrmPripaidExpenses.frx":1D9C5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1560
            TabIndex        =   62
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
            ButtonImage     =   "FrmPripaidExpenses.frx":24227
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   405
            Left            =   2760
            TabIndex        =   74
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…  Õ·Ì·Ì"
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
            ButtonImage     =   "FrmPripaidExpenses.frx":245C1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton7 
            Height          =   330
            Left            =   5880
            TabIndex        =   80
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”Œ… „„«À·…"
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
            ButtonImage     =   "FrmPripaidExpenses.frx":2AE23
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   9600
         TabIndex        =   50
         Top             =   600
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton ISButton4 
         Height          =   330
         Left            =   5880
         TabIndex        =   51
         ToolTipText     =   "Õ–› «·ﬂ·"
         Top             =   600
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
         ButtonImage     =   "FrmPripaidExpenses.frx":31685
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   570
         Left            =   120
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   0
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
      Begin ImpulseButton.ISButton ISButton6 
         Height          =   330
         Left            =   7560
         TabIndex        =   73
         ToolTipText     =   "Õ–› «·ﬂ·"
         Top             =   600
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·’› «·„Õœœ"
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
         ButtonImage     =   "FrmPripaidExpenses.frx":37EE7
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   12840
         TabIndex        =   52
         Top             =   600
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   21000
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
            Picture         =   "FrmPripaidExpenses.frx":3E749
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPripaidExpenses.frx":3EAE3
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPripaidExpenses.frx":3EE7D
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPripaidExpenses.frx":3F217
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPripaidExpenses.frx":3F5B1
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPripaidExpenses.frx":3F94B
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPripaidExpenses.frx":3FCE5
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPripaidExpenses.frx":4027F
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   21000
      TabIndex        =   55
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
      ButtonImage     =   "FrmPripaidExpenses.frx":40619
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   19560
      TabIndex        =   58
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
      ButtonImage     =   "FrmPripaidExpenses.frx":46E7B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   22320
      TabIndex        =   59
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
      ButtonImage     =   "FrmPripaidExpenses.frx":4D6DD
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
      Left            =   20880
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmPripaidExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim II As Long
 Public LonRow As Double
Public LngCol As Double
Dim flagX As Boolean

Private Sub ChkALL_Click()
    Dim I As Integer

    If ChkAll.value = vbChecked Then

        With Me.Grid
        
 
            For I = 1 To .Rows - 1
        
                .TextMatrix(I, .ColIndex("ch")) = True
            Next I

        End With

    Else

        With Me.Grid

            For I = 1 To .Rows - 1
        
                .TextMatrix(I, .ColIndex("ch")) = False
            Next I

        End With
         
    End If
    
End Sub

Private Sub DBCboClientName_Change()
DBCboClientName_Click (0)
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
TxtCustCode.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DBCboClientName.BoundText)
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 29121

    End If
End Sub

Private Sub DcbAccount_Click(Area As Integer)

TxtAccount.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
End Sub

Private Sub DcbAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 29122

    End If
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
EmpCode = GetAccountEmployee(val(DcboEmpName.BoundText))
DBCboClientName.BoundText = EmpCode
EmpCode = ""
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
End Sub
Function GetAccountEmployee(Optional EmID As Integer = 0) As String
Dim Rs7 As ADODB.Recordset
Dim sql As String
If EmID <> 0 Then
sql = "Select Account_Code3 from TblEmployee where Emp_ID =" & EmID & " "
Set Rs7 = New ADODB.Recordset
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
GetAccountEmployee = IIf(IsNull(Rs7("Account_Code3").value), " ", Rs7("Account_Code3").value)
Else
GetAccountEmployee
End If
End If

End Function
Sub maxx(Optional ByRef ID As Double = 0, Optional ByRef IDDet As Double = 0)
     
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
  Set RsDev = New ADODB.Recordset
    If ID <> 0 Then
   StrSQL = " select max(ID) as mx from ExpensesSearial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   ID = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "ExpensesSearial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("ID").value = ID
RsDev.update
End If
    If IDDet <> 0 Then
   StrSQL = " select max(IDDet) as mx from ExpensesSearial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   IDDet = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "ExpensesSearial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("IDDet").value = IDDet
RsDev.update
End If
End Sub
Function Checked(Optional ID As Double = 0, Optional IDDet As Double = 0) As Boolean
     Checked = False
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
      Set RsDev = New ADODB.Recordset
    If ID <> 0 Then
   StrSQL = " select * from ExpensesSearial where ID=" & ID & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
    If IDDet <> 0 Then
  StrSQL = " select * from ExpensesSearial where IDDet=" & IDDet & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
End Function

Private Sub dcproject_KeyUp(KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 34
             FrmProjectSearch.show vbModal
           
        End If
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
       If SystemOptions.UserInterface = ArabicInterface Then
                Grid.ColComboList(Grid.ColIndex("TypeExpens")) = "#1;  Õ”«»|#2; „ÊŸ›"
                Grid.ColComboList(Grid.ColIndex("Distribution")) = "#1;  ÌœÊÌ|#2; «·Ì"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               Grid.ColComboList(Grid.ColIndex("TypeExpens")) = "#1;Account  |#2;Eployee "
                Grid.ColComboList(Grid.ColIndex("Distribution")) = "#1;Manual  |#2;Auto "
            End If
            
      flagX = False
      
    conection = "select * from TblPripaidExpenses order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
   'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCboUserName, My_SQL
    Dim Dcombos As New ClsDataCombos
    
         
Dcombos.GetCostCenter DcCostCenter
Dcombos.GetProjects dcproject

    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetAccountingCodes Me.DcbAccount, True, False
   ' Dcombos.GetAccountingCodes Me.DcbAccount
      Dcombos.GetAccountingCodes Me.DBCboClientName, True, False
     ' Dcombos.GetAccountingCodes Me.DBCboClientName
    Dcombos.GetEmployees Me.DcboEmpName
   ' Dcombos.GetEmpDepartments Me.DcbDepartment
    BtnLast_Click
DcboEmpName.Enabled = False
DBCboClientName.Enabled = False
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
ErrTrap:
End Sub

' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
   
    If TxtModFlg = "E" Then
    RemoveGridRow4
   ' StrSQL = "Delete From TblPripaidExpensesDet Where PaidExID='" & val(TxtSerial1.text) & "'"
   ' Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    RsSavRec.Fields("RecordM").value = XPDtbTrans.value
    RsSavRec.Fields("Remark").value = Me.TxtRemark.Text
    RsSavRec.Fields("Name").value = Me.TxtName.Text
    RsSavRec.Fields("NameE").value = Me.TxtNameE.Text
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("EmpID").value = val(Me.DcboEmpName.BoundText)
    RsSavRec.Fields("Account_Code").value = Me.DBCboClientName.BoundText
    RsSavRec.Fields("Account_Code1").value = Me.DcbAccount.BoundText
    RsSavRec.Fields("Valu").value = val(Me.TxtValu.Text)
    RsSavRec.Fields("HistoryDate").value = HistoryDate.value
    RsSavRec.Fields("FromDate").value = FromDate.value
    RsSavRec.Fields("ToDate").value = ToDate.value
    RsSavRec.Fields("Remark2").value = Me.TxtRemark2.Text
    ''''''''''''''''''''''''''''''''''''''''''''''''''
     If ChMessier.value = vbChecked Then
    RsSavRec.Fields("Messier").value = 1
    Else
    RsSavRec.Fields("Messier").value = 0
    End If
    If Opt(0).value = True Then
    RsSavRec.Fields("TypeExpens").value = 0
    ElseIf Opt(1).value = True Then
    RsSavRec.Fields("TypeExpens").value = 1
     Else
    RsSavRec.Fields("TypeExpens").value = Null
    End If
    
    
    
    
    If Opttype(0).value = True Then
    RsSavRec.Fields("NewOrOpeneing").value = 0
    ElseIf Opt(1).value = True Then
    RsSavRec.Fields("NewOrOpeneing").value = 1
     Else
    RsSavRec.Fields("NewOrOpeneing").value = Null
    End If
    
     
     
    ''/////
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    ' save grid
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblPripaidExpensesDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim I As Integer
    Dim str2 As String
    With Grid
       For I = 1 To .Rows - 1
     If val(.TextMatrix(I, .ColIndex("BranchID"))) <> 0 Then
              
        ID = 0
        
        ID = IIf(Not IsNumeric(.TextMatrix(I, .ColIndex("id"))), 0, .TextMatrix(I, .ColIndex("id")))
        
          If Me.Checked(ID, 0) = True Then
        Else
       ID = 1
        maxx ID, 0
        End If
              .TextMatrix(I, .ColIndex("id")) = ID
     If ChekExpens(ID) = False Then
       RsDevsub.AddNew
                RsDevsub("ID").value = IIf(Not IsNumeric(.TextMatrix(I, .ColIndex("id"))), 0, .TextMatrix(I, .ColIndex("id")))
                RsDevsub("PaidExID").value = Me.TxtSerial1.Text
                RsDevsub("Name").value = IIf((.TextMatrix(I, .ColIndex("Name"))) = "", Null, .TextMatrix(I, .ColIndex("Name")))
                
                RsDevsub("Messier").value = IIf((.TextMatrix(I, .ColIndex("Messier"))) = "", 0, val(.TextMatrix(I, .ColIndex("Messier"))))
                RsDevsub("NameE").value = IIf((.TextMatrix(I, .ColIndex("NameE"))) = "", Null, .TextMatrix(I, .ColIndex("NameE")))
                RsDevsub("BranchID").value = IIf((.TextMatrix(I, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(I, .ColIndex("BranchID"))))
                RsDevsub("TypeExpens").value = IIf((.TextMatrix(I, .ColIndex("TypeExpens"))) = "", Null, val(.TextMatrix(I, .ColIndex("TypeExpens"))))
                
                     RsDevsub("NewOrOpeneing").value = IIf((.TextMatrix(I, .ColIndex("NewOrOpeneing"))) = "", Null, val(.TextMatrix(I, .ColIndex("NewOrOpeneing"))))
                     
                RsDevsub("EmpID").value = IIf((.TextMatrix(I, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(I, .ColIndex("EmpID"))))
                RsDevsub("Account_Code").value = IIf((.TextMatrix(I, .ColIndex("Account_Code"))) = "", Null, .TextMatrix(I, .ColIndex("Account_Code")))
                RsDevsub("Account_Code1").value = IIf((.TextMatrix(I, .ColIndex("Account_Code1"))) = "", Null, .TextMatrix(I, .ColIndex("Account_Code1")))
                RsDevsub("HistoryDate").value = IIf((.TextMatrix(I, .ColIndex("HistoryDate"))) = "", Null, .TextMatrix(I, .ColIndex("HistoryDate")))
                RsDevsub("FromDate").value = IIf((.TextMatrix(I, .ColIndex("FromDate"))) = "", Null, .TextMatrix(I, .ColIndex("FromDate")))
                RsDevsub("ToDate").value = IIf((.TextMatrix(I, .ColIndex("ToDate"))) = "", Null, .TextMatrix(I, .ColIndex("ToDate")))
                RsDevsub("Valu").value = IIf((.TextMatrix(I, .ColIndex("Valu"))) = "", Null, .TextMatrix(I, .ColIndex("Valu")))
                RsDevsub("Remark2").value = IIf((.TextMatrix(I, .ColIndex("Remark2"))) = "", Null, .TextMatrix(I, .ColIndex("Remark2")))
                
                RsDevsub("CostCenterID").value = IIf((.TextMatrix(I, .ColIndex("CostCenterID"))) = "", Null, .TextMatrix(I, .ColIndex("CostCenterID")))
                RsDevsub("CostCenterIDName").value = IIf((.TextMatrix(I, .ColIndex("CostCenterIDName"))) = "", Null, .TextMatrix(I, .ColIndex("CostCenterIDName")))
                
                RsDevsub("ProjectID").value = IIf((.TextMatrix(I, .ColIndex("ProjectID"))) = "", Null, .TextMatrix(I, .ColIndex("ProjectID")))
                RsDevsub("ProjectName").value = IIf((.TextMatrix(I, .ColIndex("ProjectName"))) = "", Null, .TextMatrix(I, .ColIndex("ProjectName")))
                
                
                RsDevsub("Distribution").value = IIf((.TextMatrix(I, .ColIndex("Distribution"))) = "", Null, val(.TextMatrix(I, .ColIndex("Distribution"))))
                 If Grid.TextMatrix(I, Grid.ColIndex("StrDistribution")) = "" Then
                   RetrStrEstam str2, I
                    .TextMatrix(I, .ColIndex("StrDistribution")) = str2
                   End If
                RsDevsub("StrDistribution").value = IIf((.TextMatrix(I, .ColIndex("StrDistribution"))) = "", Null, .TextMatrix(I, .ColIndex("StrDistribution")))
                 
                   
                   If Me.TxtModFlg.Text = "E" Then
                          StrSQL = "Delete From TblPripaidExpChiled Where PaidExIDDet =" & val(.TextMatrix(I, .ColIndex("id"))) & ""
            Cn.Execute StrSQL, , adExecuteNoRecords
                  End If
      RsDevsub.update
      saveDetails I, RsDevsub("id").value
      End If
      End If
     Next I
    End With
    flagX = False
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
               Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
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
   Sub RetrStrEstam(Optional ByRef str1 As String, Optional Row As Integer)
Dim str As String
Dim Diff As Integer
Dim StrtDate As Date
Dim cunt As Integer
Dim IDDet As Double
Dim SumVal As Double
cunt = 1
SumVal = 0
  With Grid
  
If .TextMatrix(Row, .ColIndex("FromDate")) <> "" And .TextMatrix(Row, .ColIndex("ToDate")) <> "" Then
Diff = DateDiff("m", .TextMatrix(Row, .ColIndex("FromDate")), .TextMatrix(Row, .ColIndex("ToDate"))) + 1
 SumVal = Round(val(.TextMatrix(Row, .ColIndex("Valu"))) / Diff, 2) * Diff - val(.TextMatrix(Row, .ColIndex("Valu")))
StrtDate = .TextMatrix(Row, .ColIndex("FromDate"))
Do While cunt <= Diff
  str = str & StrtDate & "#"
  If cunt = Diff And flagX = True Then
  str = str & Round(val(.TextMatrix(Row, .ColIndex("Valu"))) / Diff, 2) - SumVal & "#"
  Else
  str = str & Round(val(.TextMatrix(Row, .ColIndex("Valu"))) / Diff, 2) & "#"
  End If
    IDDet = 1
       maxx 0, IDDet
  str = str & " " & "#"
  str = str & IDDet & "#"
  
  StrtDate = DateAdd("m", 1, StrtDate)
   str = str & Trim("@")
  str = str & CHR(13)
  cunt = cunt + 1
Loop

  str1 = Trim(str)
End If
  End With
End Sub

' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim I As Integer
    ProgressBar1.Visible = True
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value): ProgressBar1.value = 10
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordM").value), Date, RsSavRec.Fields("RecordM").value): ProgressBar1.value = 20
    TxtRemark.Text = IIf(IsNull(RsSavRec.Fields("Remark").value), "", RsSavRec.Fields("Remark").value): ProgressBar1.value = 30
    TxtName.Text = IIf(IsNull(RsSavRec.Fields("Name").value), "", RsSavRec.Fields("Name").value): ProgressBar1.value = 40
    TxtNameE.Text = IIf(IsNull(RsSavRec.Fields("NameE").value), "", RsSavRec.Fields("NameE").value): ProgressBar1.value = 50
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value): ProgressBar1.value = 60
    DcboEmpName.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value): ProgressBar1.value = 70
    DBCboClientName.BoundText = IIf(IsNull(RsSavRec.Fields("Account_Code").value), "", RsSavRec.Fields("Account_Code").value): ProgressBar1.value = 80
    TxtValu.Text = IIf(IsNull(RsSavRec.Fields("Valu").value), "", RsSavRec.Fields("Valu").value): ProgressBar1.value = 90
    TxtRemark2.Text = IIf(IsNull(RsSavRec.Fields("Remark2").value), "", RsSavRec.Fields("Remark2").value): ProgressBar1.value = 100
    HistoryDate.value = IIf(IsNull(RsSavRec.Fields("HistoryDate").value), Date, RsSavRec.Fields("HistoryDate").value): ProgressBar1.value = 10
    FromDate.value = IIf(IsNull(RsSavRec.Fields("FromDate").value), Date, RsSavRec.Fields("FromDate").value): ProgressBar1.value = 20
    ToDate.value = IIf(IsNull(RsSavRec.Fields("ToDate").value), Date, RsSavRec.Fields("ToDate").value): ProgressBar1.value = 30
If (IsNull(RsSavRec.Fields("TypeExpens").value)) Then
ElseIf val(RsSavRec.Fields("TypeExpens").value) = 0 Then
Opt(0).value = True: ProgressBar1.value = 40
 Opt_Click (0)
ElseIf val(RsSavRec.Fields("TypeExpens").value) = 1 Then
Opt(1).value = True: ProgressBar1.value = 50
Opt_Click (0)
End If


If (IsNull(RsSavRec.Fields("NewOrOpeneing").value)) Then
Opttype(0).value = True
ElseIf val(RsSavRec.Fields("NewOrOpeneing").value) = 0 Then
Opttype(0).value = True
 
ElseIf val(RsSavRec.Fields("NewOrOpeneing").value) = 1 Then
Opttype(1).value = True
 
End If


If (RsSavRec.Fields("Messier").value) = True Then
ChMessier.value = vbChecked
Else
ChMessier.value = vbUnchecked
End If
 Me.DcbAccount.BoundText = IIf(IsNull(RsSavRec.Fields("Account_Code1").value), "", RsSavRec.Fields("Account_Code1").value): ProgressBar1.value = 60
     ''''''''''''''''
     DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value): ProgressBar1.value = 70

    
     LabCurrRec.Caption = RsSavRec.AbsolutePosition: ProgressBar1.value = 80
     LabCountRec.Caption = RsSavRec.RecordCount: ProgressBar1.value = 90
     ' grid
    FullGrid
 ProgressBar1.Visible = False
 ProgressBar1.value = 0
ErrTrap:
  ProgressBar1.Visible = False
 ProgressBar1.value = 0
End Sub
  Sub FullGrid()
    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
   Dim sql As String
   sql = "SELECT  ProjectID, ProjectName, TblPripaidExpensesDet.CostCenterIDName,TblPripaidExpensesDet.CostCenterID , TblPripaidExpensesDet.NewOrOpeneing ,   dbo.TblPripaidExpensesDet.ID, dbo.TblPripaidExpensesDet.Name, dbo.TblPripaidExpensesDet.NameE, dbo.TblPripaidExpensesDet.BranchID, "
   sql = sql & "                   dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblPripaidExpensesDet.TypeExpens, dbo.TblPripaidExpensesDet.EmpID,"
   sql = sql & "                    dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblPripaidExpensesDet.Account_Code, dbo.ACCOUNTS.Account_Name,"
   sql = sql & "                    dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.TblPripaidExpensesDet.HistoryDate, dbo.TblPripaidExpensesDet.FromDate,"
   sql = sql & "                    dbo.TblPripaidExpensesDet.ToDate, dbo.TblPripaidExpensesDet.Valu, dbo.TblPripaidExpensesDet.Remark2, dbo.TblPripaidExpensesDet.Distribution,"
   sql = sql & "                    dbo.TblPripaidExpensesDet.StrDistribution, dbo.TblPripaidExpensesDet.PaidExID, dbo.TblPripaidExpensesDet.Account_Code1,"
   sql = sql & "                    ACCOUNTS_1.Account_Name AS Account_Name1, ACCOUNTS_1.Account_Serial AS Account_Serial1, ACCOUNTS_1.Account_NameEng AS Account_NameE1 , dbo.TblPripaidExpensesDet.Messier"
   sql = sql & "   FROM         dbo.TblPripaidExpensesDet LEFT OUTER JOIN"
   sql = sql & "                    dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TblPripaidExpensesDet.Account_Code1 = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
   sql = sql & "                    dbo.ACCOUNTS ON dbo.TblPripaidExpensesDet.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
   sql = sql & "                    dbo.TblEmployee ON dbo.TblPripaidExpensesDet.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
   sql = sql & "                    dbo.TblBranchesData ON dbo.TblPripaidExpensesDet.BranchID = dbo.TblBranchesData.branch_id"

   sql = sql & " Where (dbo.TblPripaidExpensesDet.PaidExID =" & val(TxtSerial1.Text) & ")"
  

   Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim I As Integer
       With Me.Grid
                    For I = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(I, .ColIndex("Ser")) = I
                   .TextMatrix(I, .ColIndex("id")) = IIf(IsNull(Rs1("ID").value), "", Rs1("ID").value)
                   .TextMatrix(I, .ColIndex("Messier")) = IIf(IsNull(Rs1("Messier").value), 0, Rs1("Messier").value)
                   .TextMatrix(I, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), "", Rs1("EmpID").value)
                   .TextMatrix(I, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   .TextMatrix(I, .ColIndex("NameE")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                   .TextMatrix(I, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), "", Rs1("BranchID").value)
                   .TextMatrix(I, .ColIndex("TypeExpens")) = IIf(IsNull(Rs1("TypeExpens").value), 0, Rs1("TypeExpens").value)
                   .TextMatrix(I, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(I, .ColIndex("HistoryDate")) = IIf(IsNull(Rs1("HistoryDate").value), "", Rs1("HistoryDate").value)
                   .TextMatrix(I, .ColIndex("FromDate")) = IIf(IsNull(Rs1("FromDate").value), "", Rs1("FromDate").value)
                   .TextMatrix(I, .ColIndex("ToDate")) = IIf(IsNull(Rs1("ToDate").value), "", Rs1("ToDate").value)
                   .TextMatrix(I, .ColIndex("Account_Code1")) = IIf(IsNull(Rs1("Account_Code1").value), "", Rs1("Account_Code1").value)
                   .TextMatrix(I, .ColIndex("Account_Serial1")) = IIf(IsNull(Rs1("Account_Serial1").value), "", Rs1("Account_Serial1").value)
                   
                   .TextMatrix(I, .ColIndex("CostCenterID")) = IIf(IsNull(Rs1("CostCenterID").value), "", Rs1("CostCenterID").value)
                   
                   .TextMatrix(I, .ColIndex("CostCenterIDName")) = IIf(IsNull(Rs1("CostCenterIDName").value), "", Rs1("CostCenterIDName").value)
                   
                  .TextMatrix(I, .ColIndex("ProjectID")) = IIf(IsNull(Rs1("ProjectID").value), "", Rs1("ProjectID").value)
                   
                   .TextMatrix(I, .ColIndex("ProjectName")) = IIf(IsNull(Rs1("ProjectName").value), "", Rs1("ProjectName").value)
                         
                         
                   
                   .TextMatrix(I, .ColIndex("NewOrOpeneing")) = IIf(IsNull(Rs1("NewOrOpeneing").value), "", Rs1("NewOrOpeneing").value)
                   
                  If val(.TextMatrix(I, .ColIndex("NewOrOpeneing"))) = 0 Then
                   .TextMatrix(I, .ColIndex("NewOrOpeneingN")) = "ÃœÌœ"
                  Else
                  .TextMatrix(I, .ColIndex("NewOrOpeneingN")) = "√›  «ÕÌ"
                  End If
                  
                   .TextMatrix(I, .ColIndex("Account_Code")) = IIf(IsNull(Rs1("Account_Code").value), "", Rs1("Account_Code").value)
                   .TextMatrix(I, .ColIndex("Account_Serial")) = IIf(IsNull(Rs1("Account_Serial").value), "", Rs1("Account_Serial").value)
                   .TextMatrix(I, .ColIndex("Valu")) = IIf(IsNull(Rs1("Valu").value), "", Rs1("Valu").value)
                   .TextMatrix(I, .ColIndex("Remark2")) = IIf(IsNull(Rs1("Remark2").value), "", Rs1("Remark2").value)
                   .TextMatrix(I, .ColIndex("Distribution")) = IIf(IsNull(Rs1("Distribution").value), "", Rs1("Distribution").value)
                   .TextMatrix(I, .ColIndex("StrDistribution")) = IIf(IsNull(Rs1("StrDistribution").value), "", Rs1("StrDistribution").value)
                      
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(I, .ColIndex("Account_Name1")) = IIf(IsNull(Rs1("Account_Name1").value), "", Rs1("Account_Name1").value)
                   
                   .TextMatrix(I, .ColIndex("Account_Name")) = IIf(IsNull(Rs1("Account_Name").value), "", Rs1("Account_Name").value)
                    .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   .TextMatrix(I, .ColIndex("Branch")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                   Else
                   .TextMatrix(I, .ColIndex("Account_Name1")) = IIf(IsNull(Rs1("Account_NameE1").value), "", Rs1("Account_NameE1").value)
                   
                    .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Namee").value), "", Rs1("Emp_Namee").value)
                    .TextMatrix(I, .ColIndex("Account_Name")) = IIf(IsNull(Rs1("Account_NameEng").value), "", Rs1("Account_NameEng").value)
                   .TextMatrix(I, .ColIndex("Branch")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                   
                   End If
                   .Cell(flexcpChecked, I, .ColIndex("ch")) = flexChecked
                    Rs1.MoveNext
             Next I
             .AutoSize 0, .Cols - 1, False
        End With
     
        Exit Sub
 End Sub


Sub CalculteDate(Optional Row As Long)
If Row <> 0 Then
Dim StrtDate As Date
Dim k As Integer
Dim Diff As Integer
Dim cunt As Integer
cunt = 1
With Grid
FmrExpEnseChiled.GRID1.Rows = 2
If .TextMatrix(Row, .ColIndex("FromDate")) <> "" And .TextMatrix(Row, .ColIndex("ToDate")) <> "" Then
Diff = DateDiff("M", .TextMatrix(Row, .ColIndex("FromDate")), .TextMatrix(Row, .ColIndex("ToDate"))) + 1

StrtDate = .TextMatrix(Row, .ColIndex("FromDate"))
Do While cunt <= Diff
k = FmrExpEnseChiled.GRID1.Rows - 1
FmrExpEnseChiled.GRID1.TextMatrix(k, FmrExpEnseChiled.GRID1.ColIndex("Ser")) = k
FmrExpEnseChiled.GRID1.TextMatrix(k, FmrExpEnseChiled.GRID1.ColIndex("RecDate")) = StrtDate
FmrExpEnseChiled.GRID1.TextMatrix(k, FmrExpEnseChiled.GRID1.ColIndex("val")) = Round(val(.TextMatrix(Row, .ColIndex("Valu"))) / Diff, 2)
StrtDate = DateAdd("m", 1, StrtDate)
cunt = cunt + 1
FmrExpEnseChiled.GRID1.Rows = FmrExpEnseChiled.GRID1.Rows + 1
Loop
End If
End With
End If
End Sub
Sub PreCalculteDate(Optional Row As Integer, Optional ByRef Ind As Integer = 0)
If Row <> 0 Then
Dim StrtDate As Date
Dim k As Integer
Dim Diff As Integer
Dim cunt As Integer
Dim Sm As Double
Dim IDDet As Double
Sm = 0
cunt = 1

 FmrExpEnseChiled.GRID1.Rows = 2
Diff = DateDiff("M", FromDate.value, ToDate.value) + 1
StrtDate = FromDate.value
Do While cunt <= Diff
k = FmrExpEnseChiled.GRID1.Rows - 1
FmrExpEnseChiled.GRID1.TextMatrix(k, FmrExpEnseChiled.GRID1.ColIndex("Ser")) = k
FmrExpEnseChiled.GRID1.TextMatrix(k, FmrExpEnseChiled.GRID1.ColIndex("RecDate")) = StrtDate
FmrExpEnseChiled.GRID1.TextMatrix(k, FmrExpEnseChiled.GRID1.ColIndex("val")) = Round(val(TxtValu.Text) / Diff, 2)
   IDDet = 1
    maxx 0, IDDet
 FmrExpEnseChiled.GRID1.TextMatrix(k, FmrExpEnseChiled.GRID1.ColIndex("id")) = IDDet
  
StrtDate = DateAdd("m", 1, StrtDate)
Sm = Sm + Round(val(TxtValu.Text) / Diff, 2)
cunt = cunt + 1
FmrExpEnseChiled.GRID1.Rows = FmrExpEnseChiled.GRID1.Rows + 1
Loop
If Sm <> val(Me.TxtValu.Text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Â‰«ﬂ ›—ﬁ Â··« "
Else
MsgBox "Halalas"
End If

FmrExpEnseChiled.txtTotal.Text = val(TxtValu.Text)
FmrExpEnseChiled.TxtValue.Text = Sm
LonRow = Row
Load FmrExpEnseChiled
  FmrExpEnseChiled.CmdOk.Enabled = True
 FmrExpEnseChiled.show vbModal
 Ind = 1
End If

End If
End Sub
Sub saveDetails(Optional I As Integer = 0, Optional PaidExIDDet As Integer = 0)
Dim RsDetails11 As ADODB.Recordset
 Dim IDDet As Double
Dim astrSplit2tems2() As String
Dim astrSplitItems() As String
Dim j As Integer
Dim st As String
Dim nElements As Integer
Dim k, m As Integer
Dim Diff As Integer
If PaidExIDDet <> 0 Then
Set RsDetails11 = New ADODB.Recordset
If Me.TxtModFlg.Text = "R" Then
StrSQL = "delete From TblPripaidExpChiled  where  PaidExID =" & PaidExIDDet
                   Cn.Execute StrSQL, , adExecuteNoRecords
End If
    StrSQL = "SELECT  *  from dbo.TblPripaidExpChiled Where (1 = -1)"
   RsDetails11.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

k = 0
     If Grid.TextMatrix(I, Grid.ColIndex("StrDistribution")) <> "" Then
          st = Grid.TextMatrix(I, Grid.ColIndex("StrDistribution"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
        
         For j = 0 To nElements - 1
         With Grid
     '   Diff = DateDiff("M", .TextMatrix(Row, .ColIndex("FromDate")), .TextMatrix(Row, .ColIndex("ToDate")))
         End With
          RsDetails11.AddNew
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         Diff = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
         Diff = Diff / 3
         m = 0
        For k = 0 To Diff - 1
        RsDetails11("PaidExID").value = val(TxtSerial1.Text)
         RsDetails11("PaidExIDDet").value = PaidExIDDet
         RsDetails11("RecDate").value = astrSplit2tems2(m)
         m = m + 1
         RsDetails11("Valu").value = val(astrSplit2tems2(m))
         m = m + 1
         RsDetails11("Remark").value = astrSplit2tems2(m)
         m = m + 1

     RsDetails11("ID").value = val(astrSplit2tems2(m))
  
        m = m + 1
         RsDetails11.update
      Next k
       Next j
          End If
End If

End Sub

Private Sub RemoveGridRow2()
Dim StrSQL As String
Dim StrMSG As String
Dim ID As Double
Dim I As Integer
Dim k As Integer
If Me.TxtModFlg.Text <> "R" Then

    With Me.Grid

        If Grid.Rows < 2 Then Exit Sub
        If SystemOptions.UserInterface = ArabicInterface Then
        StrMSG = "”Ê› Ì „ Õ–› ﬂ· «·⁄„·Ì«  «·„— »ÿÂ »Â–« «·Õ—ﬂ… Â·  —Ìœ «·Õ–›"
        Else
        StrMSG = "It will be deleted all operations associated with this Transaction"
        End If
        If MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title) = vbYes Then
        k = Grid.Rows - 1
        For I = .FixedRows To Grid.Rows - 1
        'If i <= val(Grid.Rows - 1) Then
       
        If .Cell(flexcpChecked, k, .ColIndex("ch")) = flexChecked Then
        ID = val(Grid.TextMatrix(k, Grid.ColIndex("id")))
        If ChekExpens(ID) = False Then
            StrSQL = "Delete From TblPripaidExpChiled Where PaidExIDDet=" & ID & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                              StrSQL = "Delete From TblPripaidExpensesDet Where ID=" & ID & ""
    Cn.Execute StrSQL, , adExecuteNoRecords
                
          ' k = k - 1
       .RemoveItem k
    
       ' End If
       Else
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "·«Ì„ﬂ‰ «·Õ–› Â‰«·ﬂ ⁄„·Ì«  „ ⁄·ﬁ… »«·”ÿ— —ﬁ„" & " " & k
       Else
       MsgBox "Can Not Delete Row Number" & " " & k
       End If
       Exit Sub
        End If
         End If
         
        k = k - 1
       
        Next I
                
        Else
        Exit Sub
        End If
    End With

    End If
End Sub
Private Sub RemoveGridRow4()
Dim StrSQL As String
Dim StrMSG As String
Dim ID As Double
Dim I As Integer
Dim k As Integer


    With Me.Grid

   
        If Grid.Rows < 2 Then Exit Sub
    
        For I = .FixedRows To Grid.Rows - 1

        ID = val(Grid.TextMatrix(I, Grid.ColIndex("id")))
        If ID <> 0 Then
        If ChekExpens(ID) = False Then
            StrSQL = "Delete From TblPripaidExpChiled Where PaidExIDDet=" & ID & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblPripaidExpensesDet Where ID=" & ID & ""
    Cn.Execute StrSQL, , adExecuteNoRecords
                
      
    
          End If
         End If

       
        Next I
     
    End With

  
End Sub
Private Sub RemoveGridRow3()
Dim StrSQL As String
Dim StrMSG As String
Dim ID As Double
Dim I As Integer
Dim k As Integer
If Me.TxtModFlg.Text <> "R" Then

    With Me.Grid

   
        If Grid.Rows < 2 Then Exit Sub
   
        k = Grid.Rows - 1
        For I = .FixedRows To Grid.Rows - 1
   
        ID = val(Grid.TextMatrix(k, Grid.ColIndex("id")))
        If ID <> 0 Then
        If ChekExpens(ID) = False Then
            StrSQL = "Delete From TblPripaidExpChiled Where PaidExIDDet=" & ID & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblPripaidExpensesDet Where ID=" & ID & ""
    Cn.Execute StrSQL, , adExecuteNoRecords

       .RemoveItem k
       Else
        If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "·«Ì„ﬂ‰ «·Õ–› Â‰«·ﬂ ⁄„·Ì«  „ ⁄·ﬁ… »«·”ÿ— —ﬁ„" & " " & k
       Else
       MsgBox "Can Not Delete Row Number" & " " & k
       End If
       Exit Sub
        End If
         End If
         
        k = k - 1
       
        Next I
     
    End With

    End If
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim rs As New ADODB.Recordset
Dim StrAccountCode As String
Dim LngRow As Long
With Grid
Select Case .ColKey(Col)

  Case "PFuLLCode"
                .TextMatrix(Row, .ColIndex("ProjectID")) = ""
                .TextMatrix(Row, .ColIndex("ProjectName")) = ""
                StrSQL = "Select expanses_account,REVENUE_account,id,Fullcode,Project_name From projects Where Fullcode='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If rs.RecordCount > 0 Then
                     .TextMatrix(Row, .ColIndex("ProjectName")) = rs!Project_name & ""
                    .TextMatrix(Row, .ColIndex("ProjectID")) = rs!ID & ""
                Else
                    .TextMatrix(Row, .ColIndex("AccountName")) = ""
                    .TextMatrix(Row, .ColIndex("PFuLLCode")) = ""
                    .TextMatrix(Row, .ColIndex("ProjectID")) = ""
                    .TextMatrix(Row, .ColIndex("ProjectName")) = ""
                End If
     

Case "ProjectName"
 StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ProjectID"), False, True)
                .TextMatrix(Row, .ColIndex("ProjectID")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ProjectName")) = .TextMatrix(Row, .ColIndex("ProjectName"))
      



 Case "Distribution"

 If val(.TextMatrix(Row, .ColIndex("Distribution"))) = 1 Then
 LonRow = Row
 Unload FmrExpEnseChiled

 If Me.TxtModFlg.Text <> "R" And Me.TxtModFlg.Text <> "R" Then
 CalculteDate Row
 End If
 Load FmrExpEnseChiled
     If FrmPripaidExpenses.TxtModFlg.Text <> "R" Then
      FmrExpEnseChiled.CmdOk.Enabled = True
      Else
     FmrExpEnseChiled.CmdOk.Enabled = False
     End If
      FmrExpEnseChiled.txtTotal.Text = val(.TextMatrix(Row, .ColIndex("Valu")))
 FmrExpEnseChiled.show vbModal
 
 End If
 End Select
End With
End Sub

Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Grid
Select Case .ColKey(Col)
Case "Show1"
If Me.TxtModFlg.Text <> "N" Then
If val(.TextMatrix(Row, .ColIndex("Distribution"))) = 1 Or val(.TextMatrix(Row, .ColIndex("Distribution"))) = 2 Then
 LonRow = Row
Load FmrExpEnseChiled
If Me.TxtModFlg.Text = "E" Then
FmrExpEnseChiled.CmdOk.Enabled = True
Else
FmrExpEnseChiled.CmdOk.Enabled = False
End If
FmrExpEnseChiled.txtTotal.Text = val(.TextMatrix(Row, .ColIndex("Valu")))
 FmrExpEnseChiled.show vbModal
End If
End If
 End Select
End With
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 35
             FrmProjectSearch.show vbModal
           
        End If
        
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  With Me.Grid
Dim rs As New ADODB.Recordset
Dim StrComboList  As String
        Select Case .ColKey(Col)
         
          Case "ProjectName"
            StrSQL = " SELECT     ID, Project_Name"
            StrSQL = StrSQL & "            From dbo.projects"
            
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "Project_Name", "ID")
                Else
                    StrComboList = Grid.BuildComboList(rs, "Project_Name", "ID")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing


     End Select
  End With
End Sub
Function ChekExpens(Optional ID As Double) As Boolean
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "SELECT     ID, Paye"
sql = sql & " FROM         dbo.TblPripaidExpensesDet"
sql = sql & " where paye=1 and id=" & ID & ""
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekExpens = True
Else
ChekExpens = False
End If
End Function
Private Sub ISButton2_Click()
If Opt(0).value = True Then
If DBCboClientName.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— «·Õ”«»"
Else
MsgBox "Please Select Account"
End If
 DBCboClientName.SetFocus
Exit Sub
End If
End If
If Opt(0).value = True Then
If DcbAccount.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— Õ”«» «·„’—Ê›"
Else
MsgBox "Please Select Account Expenses"
End If
DcbAccount.SetFocus
Exit Sub
End If
End If

If Opt(1).value = True Then
If DcboEmpName.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— «·„ÊŸ›"
Else
MsgBox "Please Select Employee"
End If
 DcboEmpName.SetFocus
Exit Sub
End If
End If
If val(Dcbranch.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  ≈Œ Ì«— «·›—⁄"
Else
MsgBox "Please Select Branch"
End If
Dcbranch.SetFocus
Exit Sub
End If

If TxtName.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· «”„ «·„’—Ê›"
Else
MsgBox "Please Enter Name of Expenses"
End If
TxtName.SetFocus
Exit Sub
End If
If val(TxtValu.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«·  «·ﬁÌ„…"
Else
MsgBox "Please Enter  Value"
End If
TxtValu.SetFocus
Exit Sub
End If
filgrid1
End Sub
Sub filgrid1()
Dim k As Integer
Dim I As Integer
Dim Ind As Integer
Ind = 1
With Grid
k = .Rows - 1
.Rows = .Rows + 1
Do While k < (.Rows - 1)
.TextMatrix(k, .ColIndex("Ser")) = k
PreCalculteDate k, Ind
.TextMatrix(k, .ColIndex("BranchID")) = val(Dcbranch.BoundText)
.TextMatrix(k, .ColIndex("Branch")) = Dcbranch.Text
.TextMatrix(k, .ColIndex("HistoryDate")) = HistoryDate.value
.TextMatrix(k, .ColIndex("name")) = TxtName.Text
If ChMessier.value = vbChecked Then
.TextMatrix(k, .ColIndex("Messier")) = 1
Else
.TextMatrix(k, .ColIndex("Messier")) = 0
End If
.TextMatrix(k, .ColIndex("nameE")) = TxtNameE.Text
If Opt(0).value = True Then
.TextMatrix(k, .ColIndex("TypeExpens")) = 1
Else
.TextMatrix(k, .ColIndex("TypeExpens")) = 2
End If


If Opttype(0).value = True Then
.TextMatrix(k, .ColIndex("NewOrOpeneing")) = 0
.TextMatrix(k, .ColIndex("NewOrOpeneingN")) = "ÃœÌœ"
Else
.TextMatrix(k, .ColIndex("NewOrOpeneing")) = 1
.TextMatrix(k, .ColIndex("NewOrOpeneingN")) = "«›  «ÕÌ"
End If


.TextMatrix(k, .ColIndex("Account_Code1")) = DcbAccount.BoundText
.TextMatrix(k, .ColIndex("Account_Serial1")) = TxtAccount.Text
.TextMatrix(k, .ColIndex("Account_Name1")) = DcbAccount.Text

.TextMatrix(k, .ColIndex("Account_Code")) = DBCboClientName.BoundText
.TextMatrix(k, .ColIndex("Account_Serial")) = TxtCustCode.Text
.TextMatrix(k, .ColIndex("Account_Name")) = DBCboClientName.Text
.TextMatrix(k, .ColIndex("EmpID")) = DcboEmpName.BoundText
.TextMatrix(k, .ColIndex("Fullcode")) = TxtSearchCode.Text
.TextMatrix(k, .ColIndex("Emp_Name")) = DcboEmpName.Text
.TextMatrix(k, .ColIndex("Valu")) = val(TxtValu.Text)
.TextMatrix(k, .ColIndex("FromDate")) = FromDate.value
.TextMatrix(k, .ColIndex("ToDate")) = ToDate.value
.TextMatrix(k, .ColIndex("Remark2")) = TxtRemark2.Text
.TextMatrix(k, .ColIndex("Distribution")) = 2

.TextMatrix(k, .ColIndex("CostCenterID")) = DcCostCenter.BoundText
.TextMatrix(k, .ColIndex("CostCenterIDName")) = DcCostCenter.Text


.TextMatrix(k, .ColIndex("ProjectID")) = dcproject.BoundText
.TextMatrix(k, .ColIndex("ProjectName")) = dcproject.Text


k = k + 1
Loop
.AutoSize 0, .Cols - 1, False
End With
End Sub
'Private Sub ISButton3_Click()
' On Error Resume Next
'    With Me.Grid
'        If .Row <= 0 Then Exit Sub
'        .RemoveItem .Row
'    End With
'End Sub


Private Sub ISButton3_Click()
print_report11
End Sub

Private Sub ISButton4_Click()
On Error Resume Next
RemoveGridRow3
'StrSQL = "Delete From TblPripaidExpChiled Where PaidExID=" & val(TxtSerial1.text) & ""
'                Cn.Execute StrSQL, , adExecuteNoRecords
'Me.Grid.Clear flexClearScrollable, flexClearEverything
'cleargriid
'   Me.Grid.Clear flexClearScrollable, flexClearEverything
'     Grid.Rows = 2
End Sub
Private Sub ISButton5_Click()
print_report
End Sub

Private Sub ISButton6_Click()
RemoveGridRow2
End Sub

Private Sub ISButton7_Click()
Dim I As Integer
    TxtModFlg.Text = "N"
    Me.DCboUserName.BoundText = user_id
    TxtSerial1.Text = ""
    XPDtbTrans.Enabled = True
   With Grid
   For I = 1 To .Rows - 1
   .TextMatrix(I, .ColIndex("id")) = 0
  .TextMatrix(I, .ColIndex("StrDistribution")) = ""
   
   Next I
   End With
   flagX = True
End Sub

Private Sub ISButton8_Click()
Load FrmPrePaidExpensesSearch
FrmPrePaidExpensesSearch.show vbModal
End Sub

Private Sub Opt_Click(Index As Integer)
 ChMessier.value = vbUnchecked
If Opt(0).value = True Then
ChMessier.Visible = False
DBCboClientName.Enabled = True
TxtCustCode.Enabled = True
DcboEmpName.Enabled = False
DcboEmpName.BoundText = 0
TxtSearchCode.Enabled = False
Else
ChMessier.Visible = True
DBCboClientName.Enabled = False
DBCboClientName.BoundText = ""
TxtCustCode.Enabled = False
DcboEmpName.Enabled = True
TxtSearchCode.Enabled = True
End If
End Sub

Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.Text)
End Sub

Private Sub TxtCustCode_Change()
DBCboClientName.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtCustCode.Text)
End Sub

Private Sub TxtCustCode_KeyPress(KeyAscii As Integer)
DBCboClientName.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtCustCode.Text)
End Sub

Private Sub txtNameE_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub
  
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
         If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
      If Dcbranch.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Dcbranch.SetFocus
            Exit Sub
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            Dcbranch.SetFocus
         End If
     End If
btnSave.Enabled = False
            '+++++++++++++++++++++++++++++++++++++++++++++++
    ' For Each CtrlTxt In Me.Controls
    '    If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
    '        If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
    '            MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
    '            CtrlTxt.SetFocus
    '            Exit Sub
    '        End If
    '    End If
    'Next
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
    btnSave.Enabled = True
    
    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblPripaidExpenses", "ID", "")
    RsSavRec.AddNew
    TxtSerial1.Text = StrRecID
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
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
     flagX = False
End Sub
' delet sub
Private Sub btnDelete_Click()
         If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim I As Integer
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
               With Grid
               For I = .FixedRows To .Rows - 1
               ID = val(.TextMatrix(I, .ColIndex("id")))
               If ID <> 0 Then
               If ChekExpens(ID) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox "·«Ì„ﬂ‰ «·Õ–› ·«‰ Â–Â «·Õ—ﬂ… „À» Â"
               Else
               MsgBox "Can Not Delete So This is Process proof"
               End If
               Exit Sub
               End If
               End If
               Next I
               End With
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
               '''''''''''''''''''''''''''''''
                'StrSQL = "Delete From TblPripaidExpensesDet Where PaidExID='" & val(TxtSerial1.text) & "'"
                ' Cn.Execute StrSQL, , adExecuteNoRecords
               '     Dim i As Integer
               '      For i = 1 To Grid.Rows - 1
               
           ' StrSQL = "Delete From TblPripaidExpChiled Where PaidExIDDet=" & val(Grid.TextMatrix(i, Grid.ColIndex("id"))) & ""
           '     Cn.Execute StrSQL, , adExecuteNoRecords
           '     Next i
                
RemoveGridRow4
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
               cleargriid
              
     End If
                            '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
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
    ISButton7.Enabled = False
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
    ISButton7.Enabled = True
        Grid.Enabled = True
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
   ISButton7.Enabled = False
   If Opt(0).value = True Then
   DcboEmpName.Enabled = False
   TxtSearchCode.Enabled = False
   DBCboClientName.Enabled = True
   TxtCustCode.Enabled = True
   Else
   DcboEmpName.Enabled = True
   TxtSearchCode.Enabled = True
   DBCboClientName.Enabled = False
   TxtCustCode.Enabled = False
   End If
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
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
         If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        Grid.Rows = Grid.Rows + 1
        Me.DCboUserName.BoundText = user_id
      '  Me.Dcbranch.BoundText = branch_id
        Frm2.Enabled = True
        Me.Dcbranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
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
    cleargriid
    Opt(0).value = True
    Opttype(0).value = True
    Opt_Click (0)
    
    TxtModFlg.Text = "N"
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = branch_id
    Dcbranch.SetFocus
     Me.Grid.Clear flexClearScrollable, flexClearEverything
     Grid.Rows = 2
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
      cleargriid
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
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
     cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
  MySQL = " SELECT     dbo.TblPripaidExpenses.ID, dbo.TblPripaidExpenses.RecordM, dbo.TblPripaidExpenses.Remark, dbo.TblPripaidExpenses.Name, dbo.TblPripaidExpenses.NameE, "
 MySQL = MySQL & "                     dbo.TblPripaidExpenses.BranchID, TblBranchesData_1.branch_name, TblBranchesData_1.branch_namee, dbo.TblPripaidExpenses.TypeExpens,"
 MySQL = MySQL & "                     dbo.TblPripaidExpenses.EmpID, TblEmployee_1.Emp_Name, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, dbo.TblPripaidExpenses.Account_Code,"
 MySQL = MySQL & "                     ACCOUNTS_2.Account_Name, ACCOUNTS_2.Account_Serial, ACCOUNTS_2.Account_NameEng, dbo.TblPripaidExpenses.HistoryDate,"
 MySQL = MySQL & "                     dbo.TblPripaidExpenses.FromDate, dbo.TblPripaidExpenses.ToDate, dbo.TblPripaidExpenses.Valu, dbo.TblPripaidExpenses.Remark2,"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet.Name AS DetName, dbo.TblPripaidExpensesDet.NameE AS DetNameE, dbo.TblPripaidExpensesDet.BranchID AS DetBranchID,"
 MySQL = MySQL & "                     TblBranchesData_1.branch_name AS Detbranch_name, TblBranchesData_1.branch_namee AS Detbranch_namee,"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet.TypeExpens AS DetTypeExpens, dbo.TblPripaidExpensesDet.EmpID AS DetEmpID, TblEmployee_1.Emp_Name AS DetEmp_Name,"
 MySQL = MySQL & "                     TblEmployee_1.Fullcode AS DetFullcode, TblEmployee_1.Emp_Namee AS DetEmp_Namee, dbo.TblPripaidExpensesDet.HistoryDate AS DetHistoryDate,"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet.FromDate AS DetFromDate, dbo.TblPripaidExpensesDet.ToDate AS DetToDate, dbo.TblPripaidExpensesDet.Valu AS DetValu,"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet.Remark2 AS DetRemark2, dbo.TblPripaidExpensesDet.Distribution, dbo.TblPripaidExpensesDet.StrDistribution,"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet.Account_Code AS DetAccount_Code, ACCOUNTS_1.Account_Name AS DetAccount_Name,"
 MySQL = MySQL & "                     ACCOUNTS_1.Account_Serial AS DetAccount_Serial, ACCOUNTS_1.Account_NameEng AS DetAccount_NameEng,"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet.Account_Code1 AS DetAccount_Code2, ACCOUNTS_3.Account_Name AS DetAccount_Name2,"
 MySQL = MySQL & "                     ACCOUNTS_3.Account_Serial AS DetAccount_Serial2, ACCOUNTS_3.Account_NameEng AS DetAccount_NameEng2"
 MySQL = MySQL & " FROM         dbo.ACCOUNTS ACCOUNTS_3 RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet ON ACCOUNTS_3.Account_Code = dbo.TblPripaidExpensesDet.Account_Code1 LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TblPripaidExpensesDet.Account_Code = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee TblEmployee_1 ON dbo.TblPripaidExpensesDet.EmpID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData TblBranchesData_1 ON dbo.TblPripaidExpensesDet.BranchID = TblBranchesData_1.branch_id RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblPripaidExpenses ON dbo.TblPripaidExpensesDet.PaidExID = dbo.TblPripaidExpenses.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.ACCOUNTS ACCOUNTS_2 ON dbo.TblPripaidExpenses.Account_Code = ACCOUNTS_2.Account_Code LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee TblEmployee_2 ON dbo.TblPripaidExpenses.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData TblBranchesData_2 ON dbo.TblPripaidExpenses.BranchID = TblBranchesData_2.branch_id"
  MySQL = MySQL & " Where (dbo.TblPripaidExpenses.id =" & val(TxtSerial1.Text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPrePaidExpenses.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPrePaidExpensesE.rpt"
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
Function print_report11(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 MySQL = "SELECT     dbo.TblPripaidExpenses.ID, dbo.TblPripaidExpenses.RecordM, dbo.TblPripaidExpenses.Remark, dbo.TblPripaidExpenses.Name, dbo.TblPripaidExpenses.NameE, "
 MySQL = MySQL & "                     dbo.TblPripaidExpenses.BranchID, TblBranchesData_1.branch_name, TblBranchesData_1.branch_namee, dbo.TblPripaidExpenses.TypeExpens,"
 MySQL = MySQL & "                     dbo.TblPripaidExpenses.EmpID, TblEmployee_1.Emp_Name, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, dbo.TblPripaidExpenses.Account_Code,"
 MySQL = MySQL & "                     ACCOUNTS_2.Account_Name, ACCOUNTS_2.Account_Serial, ACCOUNTS_2.Account_NameEng, dbo.TblPripaidExpenses.HistoryDate,"
 MySQL = MySQL & "                     dbo.TblPripaidExpenses.FromDate, dbo.TblPripaidExpenses.ToDate, dbo.TblPripaidExpenses.Valu, dbo.TblPripaidExpenses.Remark2,"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet.Name AS DetName, dbo.TblPripaidExpensesDet.NameE AS DetNameE, dbo.TblPripaidExpensesDet.BranchID AS DetBranchID,"
 MySQL = MySQL & "                     TblBranchesData_1.branch_name AS Detbranch_name, TblBranchesData_1.branch_namee AS Detbranch_namee,"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet.TypeExpens AS DetTypeExpens, dbo.TblPripaidExpensesDet.EmpID AS DetEmpID, TblEmployee_1.Emp_Name AS DetEmp_Name,"
 MySQL = MySQL & "                     TblEmployee_1.Fullcode AS DetFullcode, TblEmployee_1.Emp_Namee AS DetEmp_Namee, dbo.TblPripaidExpensesDet.HistoryDate AS DetHistoryDate,"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet.FromDate AS DetFromDate, dbo.TblPripaidExpensesDet.ToDate AS DetToDate, dbo.TblPripaidExpensesDet.Valu AS DetValu,"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet.Remark2 AS DetRemark2, dbo.TblPripaidExpensesDet.Distribution, dbo.TblPripaidExpensesDet.StrDistribution,"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet.Account_Code AS DetAccount_Code, ACCOUNTS_1.Account_Name AS DetAccount_Name,"
 MySQL = MySQL & "                     ACCOUNTS_1.Account_Serial AS DetAccount_Serial, ACCOUNTS_1.Account_NameEng AS DetAccount_NameEng, dbo.TblPripaidExpensesDet.ID AS IDDet,"
 MySQL = MySQL & "                     dbo.TblPripaidExpChiled.Remark AS ChildRemark, dbo.TblPripaidExpChiled.RecDate, dbo.TblPripaidExpChiled.Valu AS ChildValu,"
 MySQL = MySQL & "                     dbo.TblPripaidExpChiled.ID AS IDChiled, dbo.TblPripaidExpensesDet.Account_Code1 AS DetAccount_Code2, ACCOUNTS_3.Account_Name AS DetAccount_Name2,"
 MySQL = MySQL & "                     ACCOUNTS_3.Account_Serial AS DetAccount_Serial2, ACCOUNTS_3.Account_NameEng AS DetAccount_NameEng2"
 MySQL = MySQL & "   FROM         dbo.ACCOUNTS ACCOUNTS_3 RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblPripaidExpensesDet ON ACCOUNTS_3.Account_Code = dbo.TblPripaidExpensesDet.Account_Code1 LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblPripaidExpChiled ON dbo.TblPripaidExpensesDet.ID = dbo.TblPripaidExpChiled.PaidExIDDet LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TblPripaidExpensesDet.Account_Code = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee TblEmployee_1 ON dbo.TblPripaidExpensesDet.EmpID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData TblBranchesData_1 ON dbo.TblPripaidExpensesDet.BranchID = TblBranchesData_1.branch_id RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblPripaidExpenses ON dbo.TblPripaidExpensesDet.PaidExID = dbo.TblPripaidExpenses.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.ACCOUNTS ACCOUNTS_2 ON dbo.TblPripaidExpenses.Account_Code = ACCOUNTS_2.Account_Code LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee TblEmployee_2 ON dbo.TblPripaidExpenses.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData TblBranchesData_2 ON dbo.TblPripaidExpenses.BranchID = TblBranchesData_2.branch_id"
 MySQL = MySQL & " Where (dbo.TblPripaidExpenses.id =" & val(TxtSerial1.Text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPrePaidExpensesDet.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPrePaidExpensesEDet.rpt"
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
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
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
Private Sub ISButton1_Click()
On Error GoTo ErrTrap
 '  If val(Me.TxtSerial1.text) <> 0 Then
 '      print_report
 '  End If
ErrTrap:
End Sub

Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Prepaid Expenses  "
    ' labell name
    ChkAll.RightToLeft = False
    ChkAll.Caption = "Select All"
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(4).Caption = "Trans ID"
    Me.lbl(2).Caption = "Date"
   lbl(12).Caption = "Remarks"
    Me.lbl(7).Caption = "Branch"
    lbl(3).Caption = "Name Arabic"
    lbl(9).Caption = "Name English"
    Frame2.Caption = "Type Expenses "
    Opt(0).Caption = "Account"
    Opt(0).RightToLeft = False
    lbl(5).Caption = "Advance Account"
    lbl(14).Caption = "Expenses Account"
     Opt(1).Caption = "Employee"
     Opt(1).RightToLeft = False
   lbl(10).Caption = "Employee"
   lbl(13).Caption = "Remarks"
   ISButton3.Caption = "Analytical Print"
    lbl(6).Caption = "From Date"
    lbl(11).Caption = "To Date"
    lbl(1).Caption = "History Date"
    lbl(0).Caption = "Value "
   ChMessier.RightToLeft = False
   ChMessier.Caption = "Visible in Payroll"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    '''''''''''''' next
ISButton2.Caption = "Add"

    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    ISButton6.Caption = "Delet Select"
    ISButton4.Caption = "Delet All"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    With Me.Grid
   .TextMatrix(0, .ColIndex("ch")) = "Select"
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Branch")) = "Branch Name"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Employee Code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
         .TextMatrix(0, .ColIndex("HistoryDate")) = "HistoryDate "
        .TextMatrix(0, .ColIndex("Show1")) = "Show"
        .TextMatrix(0, .ColIndex("name")) = "Name Arabic"
         .TextMatrix(0, .ColIndex("nameE")) = "Name English"
        .TextMatrix(0, .ColIndex("TypeExpens")) = "Type Expens"
        .TextMatrix(0, .ColIndex("Account_Serial")) = "Account Code "
        .TextMatrix(0, .ColIndex("Account_Name")) = "Advance Account"
        .TextMatrix(0, .ColIndex("Account_Name1")) = "Expenses Account"
        .TextMatrix(0, .ColIndex("Valu")) = "Value"
        .TextMatrix(0, .ColIndex("FromDate")) = "From Date "
        .TextMatrix(0, .ColIndex("ToDate")) = "To Date "
        .TextMatrix(0, .ColIndex("Distribution")) = "Distribution "
       .TextMatrix(0, .ColIndex("Remark2")) = "Remarks "
        
    End With
ErrTrap:
End Sub
Private Sub cleargriid()
Me.Grid.Rows = 1
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblPripaidExpenses"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end

