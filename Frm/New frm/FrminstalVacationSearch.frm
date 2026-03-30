VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmInstalVacationSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13455
   Icon            =   "FrminstalVacationSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   13455
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
         Caption         =   "«·»ÕÀ ⁄‰ «·«—’œ… «·«›  «ÕÌ… ··«Ã«“« "
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
         Picture         =   "FrminstalVacationSearch.frx":6852
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrminstalVacationSearch.frx":15141
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
         ButtonImage     =   "FrminstalVacationSearch.frx":1535C
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
         ButtonImage     =   "FrminstalVacationSearch.frx":1BBBE
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
         ButtonImage     =   "FrminstalVacationSearch.frx":22420
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
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3720
      Width           =   5955
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
         Left            =   2400
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
         Left            =   4440
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
         Left            =   1560
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3720
      Width           =   7455
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   5040
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   90701827
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   90701827
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateFromH 
         Height          =   330
         Left            =   3600
         TabIndex        =   38
         Top             =   240
         Width           =   1335
         _extentx        =   2355
         _extenty        =   582
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateToH 
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1335
         _extentx        =   2355
         _extenty        =   450
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·Õ—ﬂ…"
         Height          =   195
         Index           =   13
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   0
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   315
         Index           =   4
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   315
         Index           =   3
         Left            =   2880
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
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Height          =   375
         Left            =   7200
         TabIndex        =   64
         Top             =   1320
         Width           =   3015
         Begin XtremeSuiteControls.RadioButton opt2 
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   65
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
            Left            =   2040
            TabIndex        =   66
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
            Left            =   1560
            TabIndex        =   67
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
            Left            =   960
            TabIndex        =   68
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
            Left            =   240
            TabIndex        =   69
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
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2E9E9&
         Height          =   375
         Left            =   7200
         TabIndex        =   58
         Top             =   960
         Width           =   3015
         Begin XtremeSuiteControls.RadioButton opt1 
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   59
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "<"
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
         Begin XtremeSuiteControls.RadioButton opt1 
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   60
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   ">"
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
         Begin XtremeSuiteControls.RadioButton opt1 
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   61
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "="
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
         Begin XtremeSuiteControls.RadioButton opt1 
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   62
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "<="
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
         Begin XtremeSuiteControls.RadioButton opt1 
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   63
            Top             =   0
            Width           =   495
            _Version        =   786432
            _ExtentX        =   873
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   ">="
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
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   7200
         TabIndex        =   52
         Top             =   600
         Width           =   3015
         Begin XtremeSuiteControls.RadioButton opt 
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   53
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
         Begin XtremeSuiteControls.RadioButton opt 
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   54
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
         Begin XtremeSuiteControls.RadioButton opt 
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   55
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
         Begin XtremeSuiteControls.RadioButton opt 
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   56
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
         Begin XtremeSuiteControls.RadioButton opt 
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   57
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
      Begin VB.TextBox TxtAbcence 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10260
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1320
         Width           =   975
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   1215
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   960
         Width           =   6855
         Begin MSComCtl2.DTPicker FromLastDate 
            Height          =   330
            Left            =   4920
            TabIndex        =   43
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   90701827
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker ToLastDate 
            Height          =   330
            Left            =   1560
            TabIndex        =   44
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   90701827
            CurrentDate     =   38887
         End
         Begin Dynamic_Byte.NourHijriCal FromLastDateH 
            Height          =   330
            Left            =   3480
            TabIndex        =   45
            Top             =   360
            Width           =   1335
            _extentx        =   2355
            _extenty        =   582
         End
         Begin Dynamic_Byte.NourHijriCal ToLastDateH 
            Height          =   330
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   1335
            _extentx        =   2355
            _extenty        =   582
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " √—ÌŒ «Œ— „»«‘—…"
            Height          =   195
            Index           =   10
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   0
            Width           =   1425
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   315
            Index           =   7
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   360
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   315
            Index           =   0
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.TextBox TxtSalWithOut 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10260
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TxtBalanceVacation 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10260
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10260
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   120
         Width           =   6855
         Begin MSComCtl2.DTPicker FromStartDate 
            Height          =   330
            Left            =   4920
            TabIndex        =   26
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   90701827
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker ToStartDate 
            Height          =   330
            Left            =   1560
            TabIndex        =   27
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   90701827
            CurrentDate     =   38887
         End
         Begin Dynamic_Byte.NourHijriCal FromStartDateH 
            Height          =   330
            Left            =   3480
            TabIndex        =   40
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   582
         End
         Begin Dynamic_Byte.NourHijriCal ToStartDateH 
            Height          =   330
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   1335
            _extentx        =   2355
            _extenty        =   582
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   315
            Index           =   9
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   315
            Index           =   8
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " √—ÌŒ √Ê· „»«‘—…"
            Height          =   195
            Index           =   1
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   0
            Width           =   1425
         End
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Bindings        =   "FrminstalVacationSearch.frx":4C042
         Height          =   315
         Left            =   7200
         TabIndex        =   32
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·€Ì«»"
         Height          =   285
         Left            =   11400
         TabIndex        =   51
         Top             =   1320
         Width           =   1965
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈Ã«“… »œÊ‰ —« »"
         Height          =   285
         Left            =   11340
         TabIndex        =   36
         Top             =   960
         Width           =   1965
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—’Ìœ «·«Ã«“…"
         Height          =   285
         Left            =   11280
         TabIndex        =   34
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ›"
         Height          =   285
         Left            =   11280
         TabIndex        =   33
         Top             =   240
         Width           =   1965
      End
   End
End
Attribute VB_Name = "FrmInstalVacationSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As FrmGeneralFundReceipt

Private Sub DtpDateFrom_Change()
If Not IsNull(DtpDateFrom.value) Then
 DtpDateFromH.value = ToHijriDate(DtpDateFrom.value)
 End If
End Sub

Private Sub DtpDateFromH_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateFrom.value = ToGregorianDate(DtpDateFromH.value)
End Sub

Private Sub DtpDateTo_Change()
If Not IsNull(DtpDateTo.value) Then
 DtpDateToH.value = ToHijriDate(DtpDateTo.value)
 End If
End Sub

Private Sub DtpDateToH_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateTo.value = ToGregorianDate(DtpDateToH.value)
End Sub

Private Sub Fg_Click()
FrmInstalVacation.FindRec val(Fg.TextMatrix(Fg.Row, 1))
End Sub
Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
   Dcombos.GetEmployees Me.DcboEmpName
   '  Dcombos.GetBoxes Me.DcboBox
   ' Dcombos.GetSalesRepData Me.DataCombo1
      Set GrdBack = New ClsBackGroundPic
    With Me.Fg
        Set .WallPaper = GrdBack.Picture
       .AutoSize 0, .Cols - 1, False
    End With
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
     SetDtpickerDate Me.DtpDateFrom
     SetDtpickerDate Me.DtpDateTo
     SetDtpickerDate Me.FromStartDate
     SetDtpickerDate Me.ToStartDate
     SetDtpickerDate Me.FromLastDate
     SetDtpickerDate Me.ToLastDate
   End Sub
   Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
        GetData
        Case 1
        clear_all Me
          Me.DtpDateFrom.value = ""
          Me.DtpDateTo.value = ""
          Me.FromStartDate.value = ""
          Me.ToStartDate.value = ""
          Me.FromLastDate.value = ""
          Me.ToLastDate.value = ""
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
    sql = "SELECT     dbo.TblInstalVacation.RecordM, dbo.TblInstalVacation.RecordH, dbo.TblInstalVacation.ID, dbo.TblInstalVacation.TypeSelect, dbo.TblInstalVacationDet.BeginDate, "
    sql = sql & "                  dbo.TblInstalVacationDet.LastDate, dbo.TblInstalVacationDet.VacBalance, dbo.TblInstalVacationDet.VacWithoutSal, dbo.TblInstalVacationDet.Abcence,"
    sql = sql & "                   dbo.TblInstalVacationDet.BeginDateH, dbo.TblInstalVacationDet.LastDateH, dbo.TblInstalVacationDet.EmpID, dbo.TblEmployee.Emp_Name,"
    sql = sql & "                  dbo.TblEmployee.fullcode , dbo.TblEmployee.Emp_Namee"
    sql = sql & "    FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    sql = sql & "                  dbo.TblInstalVacationDet ON dbo.TblEmployee.Emp_ID = dbo.TblInstalVacationDet.EmpID RIGHT OUTER JOIN"
    sql = sql & "                  dbo.TblInstalVacation ON dbo.TblInstalVacationDet.InslVaID = dbo.TblInstalVacation.ID"
    
       BolBegine = False
       StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblInstalVacation.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstalVacation.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblInstalVacation.ID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblInstalVacation.ID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstalVacation.RecordM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstalVacation.RecordM>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstalVacation.RecordM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblInstalVacation.RecordM<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcboEmpName.text <> "" And (val(DcboEmpName.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblInstalVacationDet.EmpID =" & Me.DcboEmpName.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblInstalVacationDet.EmpID =" & Me.DcboEmpName.BoundText & ""
       End If
     End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(Me.FromStartDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.BeginDate >=" & SQLDate(Me.FromStartDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstalVacationDet.BeginDate >=" & SQLDate(Me.FromStartDate.value, True) & ""
        End If
    End If
    If Not IsNull(Me.ToStartDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.BeginDate <=" & SQLDate(Me.ToStartDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblInstalVacationDet.BeginDate <=" & SQLDate(Me.ToStartDate.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         If Not IsNull(Me.FromLastDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.LastDate >=" & SQLDate(Me.FromLastDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstalVacationDet.LastDate >=" & SQLDate(Me.FromLastDate.value, True) & ""
        End If
    End If
    If Not IsNull(Me.ToLastDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.LastDate <=" & SQLDate(Me.ToLastDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblInstalVacationDet.LastDate <=" & SQLDate(Me.ToLastDate.value, True) & ""
        End If
    End If
  '''''''''''''''''''''//////////////
    If val(Me.TxtAbcence.text) <> 0 Then
        If BolBegine = True Then
        If opt2(0).value = True Then
          StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.Abcence <" & val(Me.TxtAbcence.text) & ""
        ElseIf opt2(1).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.Abcence >" & val(Me.TxtAbcence.text) & ""
         
         ElseIf opt2(2).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.Abcence =" & val(Me.TxtAbcence.text) & ""
         ElseIf opt2(3).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.Abcence <=" & val(Me.TxtAbcence.text) & ""
         ElseIf opt2(4).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.Abcence >=" & val(Me.TxtAbcence.text) & ""
       End If
     Else
          BolBegine = True
          If opt2(0).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.Abcence <" & val(Me.TxtAbcence.text) & ""
         ElseIf opt2(1).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.Abcence >" & val(Me.TxtAbcence.text) & ""
          ElseIf opt2(2).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.Abcence =" & val(Me.TxtAbcence.text) & ""
          ElseIf opt2(3).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.Abcence <=" & val(Me.TxtAbcence.text) & ""
          ElseIf opt2(4).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.Abcence >=" & val(Me.TxtAbcence.text) & ""
        End If
       End If
    End If
  '''''''''/////////////////////////
      If val(Me.TxtSalWithOut.text) <> 0 Then
        If BolBegine = True Then
        If opt1(0).value = True Then
          StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.VacWithoutSal <" & val(Me.TxtSalWithOut.text) & ""
        ElseIf opt1(1).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.VacWithoutSal >" & val(Me.TxtSalWithOut.text) & ""
         
         ElseIf opt1(2).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.VacWithoutSal =" & val(Me.TxtSalWithOut.text) & ""
         ElseIf opt1(3).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.VacWithoutSal <=" & val(Me.TxtSalWithOut.text) & ""
         ElseIf opt1(4).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.VacWithoutSal >=" & val(Me.TxtSalWithOut.text) & ""
       End If
     Else
          BolBegine = True
          If opt1(0).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.VacWithoutSal <" & val(Me.TxtSalWithOut.text) & ""
         ElseIf opt1(1).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.VacWithoutSal >" & val(Me.TxtSalWithOut.text) & ""
          ElseIf opt1(2).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.VacWithoutSal =" & val(Me.TxtSalWithOut.text) & ""
          ElseIf opt1(3).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.VacWithoutSal <=" & val(Me.TxtSalWithOut.text) & ""
          ElseIf opt1(4).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.VacWithoutSal >=" & val(Me.TxtSalWithOut.text) & ""
        End If
       End If
    End If
  ''''''''''''/////////////////////
      If val(Me.TxtBalanceVacation.text) <> 0 Then
        If BolBegine = True Then
        If opt(0).value = True Then
          StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.VacBalance <" & val(Me.TxtBalanceVacation.text) & ""
        ElseIf opt(1).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.VacBalance >" & val(Me.TxtBalanceVacation.text) & ""
         
         ElseIf opt(2).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.VacBalance =" & val(Me.TxtBalanceVacation.text) & ""
         ElseIf opt(3).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.VacBalance <=" & val(Me.TxtBalanceVacation.text) & ""
         ElseIf opt(4).value = True Then
         StrWhere = StrWhere & " AND dbo.TblInstalVacationDet.VacBalance >=" & val(Me.TxtBalanceVacation.text) & ""
       End If
     Else
          BolBegine = True
          If opt(0).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.VacBalance <" & val(Me.TxtBalanceVacation.text) & ""
         ElseIf opt(1).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.VacBalance >" & val(Me.TxtBalanceVacation.text) & ""
          ElseIf opt(2).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.VacBalance =" & val(Me.TxtBalanceVacation.text) & ""
          ElseIf opt(3).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.VacBalance <=" & val(Me.TxtBalanceVacation.text) & ""
          ElseIf opt(4).value = True Then
         StrWhere = " Where dbo.TblInstalVacationDet.VacBalance >=" & val(Me.TxtBalanceVacation.text) & ""
        End If
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblInstalVacation.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.Fg
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
                   If Not (IsNull(rs("BeginDate").value)) Then
                .TextMatrix(i, .ColIndex("BeginDate")) = Format(rs("BeginDate").value, "yyyy/M/d")
                End If
                   If Not (IsNull(rs("LastDate").value)) Then
                .TextMatrix(i, .ColIndex("LastDate")) = Format(rs("LastDate").value, "yyyy/M/d")
                End If
                
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                Else
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
               End If
               .TextMatrix(i, .ColIndex("RecordH")) = IIf(IsNull(rs("RecordH").value), "", rs("RecordH").value)
               .TextMatrix(i, .ColIndex("LastDateH")) = IIf(IsNull(rs("LastDateH").value), "", rs("LastDateH").value)
               .TextMatrix(i, .ColIndex("BeginDateH")) = IIf(IsNull(rs("BeginDateH").value), "", rs("BeginDateH").value)
               .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)
               .TextMatrix(i, .ColIndex("VacBalance")) = IIf(IsNull(rs("VacBalance").value), "", rs("VacBalance").value)
               .TextMatrix(i, .ColIndex("VacWithoutSal")) = IIf(IsNull(rs("VacWithoutSal").value), "", rs("VacWithoutSal").value)
               .TextMatrix(i, .ColIndex("Abcence")) = IIf(IsNull(rs("Abcence").value), "", rs("Abcence").value)
               
               
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub

Private Sub ChangeLang()
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
     Me.Caption = "Opening Balances Search"
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(14).Caption = "Trans ID"
    Me.lbl(5).Caption = "From"
    Me.lbl(6).Caption = "To"
    Me.lbl(13).Caption = "Trans Date"
    Me.lbl(4).Caption = "From"
    Me.lbl(3).Caption = "To"
    Label2.Caption = "Employee"
    Label3.Caption = "Balance Vacation"
    Label4.Caption = "Unpaid Vacation"
    Label5.Caption = "Absence"
    lbl(2).Caption = "Total"
    lbl(1).Caption = "Start Date"
    lbl(10).Caption = "Last Date"
    lbl(8).Caption = "From"
    lbl(9).Caption = "To"
    lbl(7).Caption = "From"
    lbl(0).Caption = "To"
    ''''''''''''''''''''''' next

     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecordM")) = "Trans Date"
        .TextMatrix(0, .ColIndex("RecordH")) = "Trans Date"
        .TextMatrix(0, .ColIndex("fullcode")) = "Employee Code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
        .TextMatrix(0, .ColIndex("VacBalance")) = "Balance Vacation"
        .TextMatrix(0, .ColIndex("VacWithoutSal")) = "Unpaid Vacation"
        .TextMatrix(0, .ColIndex("Abcence")) = "Abcence"
        .TextMatrix(0, .ColIndex("BeginDate")) = "Start Date"
        .TextMatrix(0, .ColIndex("BeginDateH")) = "Start Date"
        .TextMatrix(0, .ColIndex("LastDate")) = "Last Date"
        .TextMatrix(0, .ColIndex("LastDateH")) = "Last Date"
       ' .TextMatrix(0, .ColIndex("Explan")) = "Explanation"
    End With
  End Sub
'''''''''''''''''''''''''''' end


Private Sub FromLastDate_Change()
If Not IsNull(FromLastDate.value) Then
 FromLastDateH.value = ToHijriDate(FromLastDate.value)
 End If
End Sub

Private Sub FromLastDateH_LostFocus()
 VBA.Calendar = vbCalGreg
            FromLastDate.value = ToGregorianDate(FromLastDateH.value)
End Sub

Private Sub FromStartDate_Change()
If Not IsNull(FromStartDate.value) Then
 FromStartDateH.value = ToHijriDate(FromStartDate.value)
 End If
End Sub

Private Sub FromStartDateH_LostFocus()
 VBA.Calendar = vbCalGreg
            FromStartDate.value = ToGregorianDate(FromStartDateH.value)
End Sub

Private Sub ToLastDate_Change()
If Not IsNull(ToLastDate.value) Then
 ToLastDateH.value = ToHijriDate(ToLastDate.value)
 End If
End Sub

Private Sub ToLastDateH_LostFocus()
 VBA.Calendar = vbCalGreg
            ToLastDate.value = ToGregorianDate(ToLastDateH.value)
End Sub

Private Sub ToStartDate_Change()
If Not IsNull(ToStartDate.value) Then
 ToStartDateH.value = ToHijriDate(ToStartDate.value)
 End If
End Sub

Private Sub ToStartDateH_LostFocus()
 VBA.Calendar = vbCalGreg
            ToStartDate.value = ToGregorianDate(ToStartDateH.value)
End Sub
