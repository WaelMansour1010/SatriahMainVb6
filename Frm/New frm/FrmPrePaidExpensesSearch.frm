VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmPrePaidExpensesSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13410
   Icon            =   "FrmPrePaidExpensesSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   13410
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   3720
      Width           =   4335
      Begin MSComCtl2.DTPicker HistroyFrom 
         Height          =   330
         Left            =   2280
         TabIndex        =   39
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95944707
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker HistroyFromTo 
         Height          =   330
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95944707
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·«À»« "
         Height          =   195
         Index           =   0
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   0
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰ "
         Height          =   315
         Index           =   1
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï "
         Height          =   315
         Index           =   8
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   0
      Width           =   13665
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·»ÕÀ ⁄‰ «·„’—Ê›«  «·„œ›Ê⁄… „ﬁœ„«"
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
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   5400
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   12360
         Picture         =   "FrmPrePaidExpensesSearch.frx":6852
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   3015
      Left            =   0
      TabIndex        =   21
      Top             =   720
      Width           =   13455
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2625
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   13155
         _cx             =   23204
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPrePaidExpensesSearch.frx":15141
         ScrollTrack     =   -1  'True
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   6360
      Width           =   13455
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   10
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·≈Ã„«·Ì"
         Height          =   285
         Index           =   2
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      TabIndex        =   13
      Top             =   6960
      Width           =   13455
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   9480
         TabIndex        =   14
         Top             =   240
         Width           =   3525
         _ExtentX        =   6218
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
         BackStyle       =   0
         ButtonImage     =   "FrmPrePaidExpensesSearch.frx":15380
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         LowerToggledContent=   0   'False
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   15
         Top             =   240
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "„”Õ"
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
         ButtonImage     =   "FrmPrePaidExpensesSearch.frx":1BBE2
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
      Begin ImpulseButton.ISButton Cmd 
         Cancel          =   -1  'True
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
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
         BackStyle       =   0
         ButtonImage     =   "FrmPrePaidExpensesSearch.frx":22444
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
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3720
      Width           =   4515
      Begin VB.TextBox TxtIDTO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Õ—ﬂ…"
         Height          =   195
         Index           =   14
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   315
         Index           =   6
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   315
         Index           =   5
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3720
      Width           =   4575
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95944707
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95944707
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·Õ—ﬂ…"
         Height          =   195
         Index           =   13
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   0
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰ "
         Height          =   315
         Index           =   4
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï "
         Height          =   315
         Index           =   3
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Height          =   1935
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4440
      Width           =   13425
      Begin VB.Frame Frame9 
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «· Ê“Ì⁄"
         Height          =   555
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   1320
         Width           =   4155
         Begin XtremeSuiteControls.RadioButton Optt 
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   67
            Top             =   240
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "ÌœÊÌ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Optt 
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   68
            Top             =   240
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«·Ì"
            ForeColor       =   0
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Optt 
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   70
            Top             =   240
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«·ﬂ·"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„’—Ê›"
         Height          =   555
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   1320
         Width           =   4275
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   64
            Top             =   240
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
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "„ÊŸ›"
            ForeColor       =   0
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   69
            Top             =   240
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«·ﬂ·"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.TextBox TxtCustCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7020
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TxtNameE 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   600
         Width           =   3555
      End
      Begin VB.TextBox TxtName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   240
         Width           =   3555
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   120
         Width           =   4575
         Begin MSComCtl2.DTPicker StrFrom 
            Height          =   330
            Left            =   2280
            TabIndex        =   51
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   95944707
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker StrTo 
            Height          =   330
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   95944707
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì»œ« „‰  «—ÌŒ"
            Height          =   195
            Index           =   16
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   0
            Width           =   1425
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ "
            Height          =   315
            Index           =   15
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï "
            Height          =   315
            Index           =   12
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   120
         Width           =   4575
         Begin MSComCtl2.DTPicker EndFrom 
            Height          =   330
            Left            =   2280
            TabIndex        =   45
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   95944707
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker EndTo 
            Height          =   330
            Left            =   240
            TabIndex        =   46
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   95944707
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï "
            Height          =   315
            Index           =   11
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ "
            Height          =   315
            Index           =   10
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì‰ ÂÌ » «—ÌŒ"
            Height          =   195
            Index           =   9
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   2775
         Begin XtremeSuiteControls.RadioButton opt2 
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   31
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "<"
            ForeColor       =   16711680
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton opt2 
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   32
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   ">"
            ForeColor       =   16711680
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton opt2 
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   33
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "="
            ForeColor       =   16711680
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton opt2 
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   34
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "<="
            ForeColor       =   16711680
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton opt2 
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   ">="
            ForeColor       =   16711680
            BackColor       =   14871017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
      End
      Begin VB.TextBox TxtValue 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7020
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Bindings        =   "FrmPrePaidExpensesSearch.frx":4C066
         Height          =   315
         Left            =   4440
         TabIndex        =   26
         Top             =   960
         Width           =   2535
         _ExtentX        =   4471
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
         Bindings        =   "FrmPrePaidExpensesSearch.frx":4C07B
         Height          =   315
         Left            =   8880
         TabIndex        =   36
         Top             =   960
         Width           =   3555
         _ExtentX        =   6271
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
      Begin MSDataListLib.DataCombo DBCboClientName 
         Height          =   315
         Left            =   4440
         TabIndex        =   61
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Õ”«»"
         Height          =   285
         Index           =   19
         Left            =   7800
         TabIndex        =   62
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ ≈‰Ã·Ì“Ì"
         Height          =   285
         Index           =   18
         Left            =   12360
         TabIndex        =   58
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ ⁄—»Ì"
         Height          =   285
         Index           =   17
         Left            =   12360
         TabIndex        =   56
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   7
         Left            =   12360
         TabIndex        =   37
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁÌ„…"
         Height          =   285
         Index           =   21
         Left            =   3720
         TabIndex        =   29
         Top             =   960
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ›"
         Height          =   285
         Index           =   20
         Left            =   7800
         TabIndex        =   27
         Top             =   1080
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FrmPrePaidExpensesSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As FrmGeneralFundReceipt

Private Sub DBCboClientName_Change()
DBCboClientName_Click (0)
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
TxtCustCode.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DBCboClientName.BoundText)
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub


    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
End Sub

Private Sub Fg_Click()

FrmPripaidExpenses.FindRec val(FG.TextMatrix(FG.Row, 1))
FrmPripaidExpenses.TxtModFlg = "R"
End Sub
Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim My_SQL As String
      If SystemOptions.UserInterface = ArabicInterface Then
                FG.ColComboList(FG.ColIndex("TypeExpens")) = "#1;  Õ”«»|#2; „ÊŸ›"
                FG.ColComboList(FG.ColIndex("Distribution")) = "#1;  ÌœÊÌ|#2; «·Ì"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               FG.ColComboList(FG.ColIndex("TypeExpens")) = "#1;Account  |#2;Eployee "
                FG.ColComboList(FG.ColIndex("Distribution")) = "#1;Manual  |#2;Auto "
            End If
            
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
   Dcombos.GetEmployees Me.DcboEmpName
   Dcombos.GetBranches Me.Dcbranch
             If SystemOptions.UserInterface = ArabicInterface Then
            My_SQL = "  select Account_Code,Account_Name from ACCOUNTS where last_account=1"
Else
My_SQL = "  select Account_Code,Account_Nameeng from ACCOUNTS where last_account=1"
End If
            fill_combo Me.DBCboClientName, My_SQL
      Set GrdBack = New ClsBackGroundPic
    With Me.FG
        Set .WallPaper = GrdBack.Picture
       .AutoSize 0, .Cols - 1, False
    End With
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
     SetDtpickerDate Me.DtpDateFrom
     SetDtpickerDate Me.DtpDateTo
     SetDtpickerDate Me.HistroyFrom
     SetDtpickerDate Me.HistroyFromTo
     SetDtpickerDate Me.StrFrom
     SetDtpickerDate Me.StrTo
      SetDtpickerDate Me.EndFrom
     SetDtpickerDate Me.EndTo
   End Sub
   Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
        GetData
        Case 1
        clear_all Me
   
          Me.DtpDateFrom.value = ""
          Me.DtpDateTo.value = ""
          Me.HistroyFrom.value = ""
          Me.HistroyFromTo.value = ""
          Me.StrFrom.value = ""
          Me.StrTo.value = ""
          Me.EndFrom.value = ""
          Me.EndTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lblL(0).Caption = "Search Results"
            End If
            Case 2
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Public Sub GetData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = "SELECT     dbo.TblPripaidExpenses.ID, dbo.TblPripaidExpenses.RecordM, dbo.TblPripaidExpenses.Remark, dbo.TblPripaidExpenses.Name, dbo.TblPripaidExpenses.NameE, "
    sql = sql & "                   dbo.TblPripaidExpenses.BranchID, TblBranchesData_1.branch_name, TblBranchesData_1.branch_namee, dbo.TblPripaidExpenses.TypeExpens,"
    sql = sql & "                  dbo.TblPripaidExpenses.EmpID, TblEmployee_1.Emp_Name, TblEmployee_1.Fullcode, TblEmployee_1.Emp_Namee, dbo.TblPripaidExpenses.Account_Code,"
    sql = sql & "                  ACCOUNTS_1.Account_Name, ACCOUNTS_1.Account_Serial, ACCOUNTS_1.Account_NameEng, dbo.TblPripaidExpenses.HistoryDate,"
    sql = sql & "                  dbo.TblPripaidExpenses.FromDate, dbo.TblPripaidExpenses.ToDate, dbo.TblPripaidExpenses.Valu, dbo.TblPripaidExpenses.Remark2,"
    sql = sql & "                  dbo.TblPripaidExpensesDet.Name AS DetName, dbo.TblPripaidExpensesDet.NameE AS DetNameE, dbo.TblPripaidExpensesDet.BranchID AS DetBranchID,"
    sql = sql & "                  TblBranchesData_1.branch_name AS Detbranch_name, TblBranchesData_1.branch_namee AS Detbranch_namee,"
    sql = sql & "                  dbo.TblPripaidExpensesDet.TypeExpens AS DetTypeExpens, dbo.TblPripaidExpensesDet.EmpID AS DetEmpID, TblEmployee_1.Emp_Name AS DetEmp_Name,"
    sql = sql & "                  TblEmployee_1.Fullcode AS DetFullcode, TblEmployee_1.Emp_Namee AS DetEmp_Namee, dbo.TblPripaidExpensesDet.HistoryDate AS DetHistoryDate,"
    sql = sql & "                  dbo.TblPripaidExpensesDet.FromDate AS DetFromDate, dbo.TblPripaidExpensesDet.ToDate AS DetToDate, dbo.TblPripaidExpensesDet.Valu AS DetValu,"
    sql = sql & "                  dbo.TblPripaidExpensesDet.Remark2 AS DetRemark2, dbo.TblPripaidExpensesDet.Distribution, dbo.TblPripaidExpensesDet.StrDistribution,"
    sql = sql & "                  dbo.TblPripaidExpensesDet.Account_Code AS DetAccount_Code, ACCOUNTS_1.Account_Name AS DetAccount_Name,"
    sql = sql & "                  ACCOUNTS_1.Account_Serial AS DetAccount_Serial, ACCOUNTS_1.Account_NameEng AS DetAccount_NameEng, dbo.TblPripaidExpensesDet.ID AS IDDet"
    sql = sql & "    FROM         dbo.TblPripaidExpensesDet LEFT OUTER JOIN"
    sql = sql & "                  dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TblPripaidExpensesDet.Account_Code = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblEmployee TblEmployee_1 ON dbo.TblPripaidExpensesDet.EmpID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData TblBranchesData_1 ON dbo.TblPripaidExpensesDet.BranchID = TblBranchesData_1.branch_id RIGHT OUTER JOIN"
    sql = sql & "                  dbo.TblPripaidExpenses ON dbo.TblPripaidExpensesDet.PaidExID = dbo.TblPripaidExpenses.ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.ACCOUNTS ACCOUNTS_2 ON dbo.TblPripaidExpenses.Account_Code = ACCOUNTS_2.Account_Code LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblEmployee TblEmployee_2 ON dbo.TblPripaidExpenses.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData TblBranchesData_2 ON dbo.TblPripaidExpenses.BranchID = TblBranchesData_2.branch_id"

    
       BolBegine = False
       StrWhere = ""
 
  If Me.Opt(0).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblPripaidExpensesDet.TypeExpens =1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblPripaidExpensesDet.TypeExpens =1"
        End If
    End If
     If Me.Opt(1).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblPripaidExpensesDet.TypeExpens =2"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblPripaidExpensesDet.TypeExpens =2"
        End If
    End If
    
        If Me.Optt(0).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblPripaidExpensesDet.Distribution =1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblPripaidExpensesDet.Distribution =1"
        End If
    End If
    
          If Me.Optt(1).value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblPripaidExpensesDet.Distribution =2"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblPripaidExpensesDet.Distribution =2"
        End If
    End If
  If Me.TxtName.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblPripaidExpensesDet.Name like '%" & Me.TxtName.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblPripaidExpensesDet.Name like '%" & Me.TxtName.Text & "%'"
        End If
    End If
      If Me.TxtNameE.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblPripaidExpensesDet.NameE like '%" & Me.TxtNameE.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblPripaidExpensesDet.NameE like '%" & Me.TxtNameE.Text & "%'"
        End If
    End If
    
    
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblPripaidExpenses.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblPripaidExpenses.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblPripaidExpenses.ID <=" & val(Me.TxtIDTO.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblPripaidExpenses.ID <=" & val(Me.TxtIDTO.Text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblPripaidExpenses.RecordM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblPripaidExpenses.RecordM>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblPripaidExpenses.RecordM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblPripaidExpenses.RecordM<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcboEmpName.Text <> "" And (val(DcboEmpName.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblPripaidExpensesDet.EmpID =" & Me.DcboEmpName.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblPripaidExpensesDet.EmpID =" & Me.DcboEmpName.BoundText & ""
       End If
     End If
        If Me.Dcbranch.Text <> "" And (val(Dcbranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblPripaidExpensesDet.BranchID =" & Me.Dcbranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblPripaidExpensesDet.BranchID =" & Me.Dcbranch.BoundText & ""
       End If
     End If
     
        If Me.DBCboClientName.Text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblPripaidExpensesDet.Account_Code ='" & Me.DBCboClientName.BoundText & "'"
        Else:
          BolBegine = True
          StrWhere = " Where TblPripaidExpensesDet.Account_Code ='" & Me.DBCboClientName.BoundText & "'"
       End If
     End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(Me.HistroyFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblPripaidExpensesDet.HistoryDate >=" & SQLDate(Me.HistroyFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblPripaidExpensesDet.HistoryDate >=" & SQLDate(Me.HistroyFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.HistroyFromTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblPripaidExpensesDet.HistoryDate <=" & SQLDate(Me.HistroyFromTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblPripaidExpensesDet.HistoryDate <=" & SQLDate(Me.HistroyFromTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         If Not IsNull(Me.StrFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblPripaidExpensesDet.FromDate >=" & SQLDate(Me.StrFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblPripaidExpensesDet.FromDate >=" & SQLDate(Me.StrFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.StrTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblPripaidExpensesDet.FromDate <=" & SQLDate(Me.StrTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblPripaidExpensesDet.FromDate <=" & SQLDate(Me.StrTo.value, True) & ""
        End If
    End If
    ''''''''''''''''''''///////
             If Not IsNull(Me.EndFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblPripaidExpensesDet.ToDate >=" & SQLDate(Me.EndFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblPripaidExpensesDet.ToDate >=" & SQLDate(Me.EndFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.EndTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblPripaidExpensesDet.ToDate <=" & SQLDate(Me.EndTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblPripaidExpensesDet.ToDate <=" & SQLDate(Me.EndTo.value, True) & ""
        End If
    End If
  '''''''''''''''''''''//////////////
    If val(Me.TxtValue.Text) <> 0 Then
        If BolBegine = True Then
        If opt2(0).value = True Then
          StrWhere = StrWhere & " AND dbo.TblPripaidExpensesDet.Valu <" & val(Me.TxtValue.Text) & ""
        ElseIf opt2(1).value = True Then
         StrWhere = StrWhere & " AND dbo.TblPripaidExpensesDet.Valu >" & val(Me.TxtValue.Text) & ""
        ElseIf opt2(2).value = True Then
         StrWhere = StrWhere & " AND dbo.TblPripaidExpensesDet.Valu =" & val(Me.TxtValue.Text) & ""
         ElseIf opt2(3).value = True Then
         StrWhere = StrWhere & " AND dbo.TblPripaidExpensesDet.Valu <=" & val(Me.TxtValue.Text) & ""
         ElseIf opt2(4).value = True Then
         StrWhere = StrWhere & " AND dbo.TblPripaidExpensesDet.Valu >=" & val(Me.TxtValue.Text) & ""
       End If
     Else
          BolBegine = True
          If opt2(0).value = True Then
         StrWhere = " Where dbo.TblPripaidExpensesDet.Valu <" & val(Me.TxtValue.Text) & ""
         ElseIf opt2(1).value = True Then
         StrWhere = " Where dbo.TblPripaidExpensesDet.Valu >" & val(Me.TxtValue.Text) & ""
          ElseIf opt2(2).value = True Then
         StrWhere = " Where dbo.TblPripaidExpensesDet.Valu =" & val(Me.TxtValue.Text) & ""
          ElseIf opt2(3).value = True Then
         StrWhere = " Where dbo.TblPripaidExpensesDet.Valu <=" & val(Me.TxtValue.Text) & ""
          ElseIf opt2(4).value = True Then
         StrWhere = " Where dbo.TblPripaidExpensesDet.Valu >=" & val(Me.TxtValue.Text) & ""
        End If
       End If
    End If

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblPripaidExpenses.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordM").value)) Then
                .TextMatrix(i, .ColIndex("RecordM")) = Format(rs("RecordM").value, "yyyy/M/d")
                End If
                   If Not (IsNull(rs("DetHistoryDate").value)) Then
                .TextMatrix(i, .ColIndex("DetHistoryDate")) = Format(rs("DetHistoryDate").value, "yyyy/M/d")
                End If
                   If Not (IsNull(rs("DetFromDate").value)) Then
                .TextMatrix(i, .ColIndex("DetFromDate")) = Format(rs("DetFromDate").value, "yyyy/M/d")
                End If
                    If Not (IsNull(rs("DetToDate").value)) Then
                .TextMatrix(i, .ColIndex("DetToDate")) = Format(rs("DetToDate").value, "yyyy/M/d")
                End If
                
                .TextMatrix(i, .ColIndex("DetName")) = IIf(IsNull(rs("DetName").value), "", rs("DetName").value)
                .TextMatrix(i, .ColIndex("DetNameE")) = IIf(IsNull(rs("DetNameE").value), "", rs("DetNameE").value)
                .TextMatrix(i, .ColIndex("TypeExpens")) = IIf(IsNull(rs("DetTypeExpens").value), 0, rs("DetTypeExpens").value)
                .TextMatrix(i, .ColIndex("Distribution")) = IIf(IsNull(rs("Distribution").value), "", rs("Distribution").value)
                 .TextMatrix(i, .ColIndex("DetValu")) = IIf(IsNull(rs("DetValu").value), 0, rs("DetValu").value)
                
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Detbranch_name")) = IIf(IsNull(rs("Detbranch_name").value), "", rs("Detbranch_name").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("DetEmp_Name").value), "", rs("DetEmp_Name").value)
                .TextMatrix(i, .ColIndex("DetAccount_Name")) = IIf(IsNull(rs("DetAccount_Name").value), "", rs("DetAccount_Name").value)
                Else
                .TextMatrix(i, .ColIndex("Detbranch_name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("DetEmp_Namee").value), "", rs("DetEmp_Namee").value)
                .TextMatrix(i, .ColIndex("DetAccount_Name")) = IIf(IsNull(rs("DetAccount_NameEng").value), "", rs("DetAccount_NameEng").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub

Private Sub ChangeLang()
    Cmd(1).Caption = "Clear"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
     Me.Caption = "PrePaidExpenses Search"
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(14).Caption = "Trans ID"
    Me.lbl(5).Caption = "From"
    Me.lbl(6).Caption = "To"
    Me.lbl(13).Caption = "Trans Date"
    Me.lbl(4).Caption = "From"
    Me.lbl(3).Caption = "To"
    lbl(0).Caption = "History Date"
    Me.lbl(1).Caption = "From"
    Me.lbl(8).Caption = "To"
    lbl(17).Caption = "Name Ar"
    lbl(18).Caption = "Name Eng"
    lbl(9).Caption = "End Date"
    Me.lbl(10).Caption = "From"
    Me.lbl(11).Caption = "To"
      lbl(16).Caption = "Start Date"
    Me.lbl(15).Caption = "From"
    Me.lbl(12).Caption = "To"
    lbl(7).Caption = "Branch"
    Frame7.Caption = "Type Expenses"
    Opt(0).Caption = "Account"
    Opt(0).RightToLeft = False
    Opt(1).Caption = "Employee"
    Opt(1).RightToLeft = False
    Opt(2).Caption = "All"
    Opt(2).RightToLeft = False
    lbl(20).Caption = "Employee"
    lbl(19).Caption = "Account"
    lbl(21).Caption = "Value"
    Frame9.Caption = "Distribution"
    lbl(2).Caption = "Total"
   Optt(0).Caption = "Manual"
   Optt(0).RightToLeft = False
   Optt(1).Caption = "Auto"
   Optt(1).RightToLeft = False
   Optt(2).Caption = "All"
   Optt(2).RightToLeft = False
    ''''''''''''''''''''''' next

     With Me.FG
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecordM")) = "Trans Date"
        .TextMatrix(0, .ColIndex("DetHistoryDate")) = "History Date"
        .TextMatrix(0, .ColIndex("DetName")) = "Name Arabic"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
        .TextMatrix(0, .ColIndex("DetNameE")) = "Name English"
        .TextMatrix(0, .ColIndex("Detbranch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("TypeExpens")) = "Type Expens"
        .TextMatrix(0, .ColIndex("DetValu")) = "Value"
        .TextMatrix(0, .ColIndex("DetAccount_Name")) = "Account"
        .TextMatrix(0, .ColIndex("Distribution")) = "Distribution "
        .TextMatrix(0, .ColIndex("DetFromDate")) = "From Date"
        .TextMatrix(0, .ColIndex("DetToDate")) = "To Date"
    End With
  End Sub
'''''''''''''''''''''''''''' end




Private Sub TxtCustCode_KeyPress(KeyAscii As Integer)
DBCboClientName.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtCustCode.Text)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub
