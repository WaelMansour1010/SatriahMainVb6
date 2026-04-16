VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmUnitRent 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ⁄ﬁœ ÃœÌœ"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15165
   Icon            =   "FrmUnitRent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   15165
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2E9E9&
      Height          =   1335
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   79
      Top             =   5040
      Width           =   15135
      Begin VB.VScrollBar VScroll6 
         Height          =   375
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   110
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   720
         Width           =   2235
      End
      Begin VB.TextBox Text23 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   240
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘»ﬂ…"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   375
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   720
         Width           =   1275
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   375
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   375
         Left            =   13080
         RightToLeft     =   -1  'True
         TabIndex        =   98
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   720
         Width           =   915
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   375
         Left            =   13080
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   240
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   13320
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   720
         Width           =   3855
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ﬁœ«"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   240
         Width           =   615
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   13
         Left            =   120
         TabIndex        =   84
         Top             =   3120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
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
         ButtonImage     =   "FrmUnitRent.frx":038A
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÌÊ„"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   15
         Left            =   11400
         TabIndex        =   99
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”«⁄…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   21
         Left            =   13440
         TabIndex        =   94
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”«⁄…"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   17
         Left            =   13440
         TabIndex        =   92
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Êﬁ  «·ﬁœÊ„"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   19
         Left            =   14040
         TabIndex        =   91
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÌÊ„"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   18
         Left            =   11400
         TabIndex        =   90
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Â«Ì… «·«ÌÃ«—"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   14
         Left            =   14040
         TabIndex        =   89
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·œ›⁄… «·„ﬁœ„…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   12
         Left            =   7320
         TabIndex        =   88
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·«Ì«„"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   10
         Left            =   7440
         TabIndex        =   87
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„·«ÕŸ«  ⁄‰ «·⁄ﬁœ"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   9
         Left            =   3960
         TabIndex        =   86
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”‰œ"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   85
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   4320
      Width           =   15135
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text25 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text24 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   78
         Top             =   3120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
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
         ButtonImage     =   "FrmUnitRent.frx":0924
         DrawFocusRectangle=   0   'False
      End
      Begin MSDataListLib.DataCombo DataCombo5 
         Height          =   315
         Left            =   12720
         TabIndex        =   117
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÊÕœ…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   22
         Left            =   13440
         TabIndex        =   118
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·√Ã— «·ÌÊ„Ì"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   16
         Left            =   3120
         TabIndex        =   116
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·œÊ—"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   13
         Left            =   7920
         TabIndex        =   114
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·ÊÕœ…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   11
         Left            =   11880
         TabIndex        =   112
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Height          =   1575
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   2760
      Width           =   15135
      Begin VB.CommandButton Command3 
         Caption         =   "Õ‹‹‹‹‹‹‹‹‹–›"
         Height          =   375
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   " ⁄‹‹‹‹‹‹‹‹‹œÌ·"
         Height          =   375
         Left            =   11280
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "≈÷‹‹‹‹‹‹‹‹‹‹‹«›…"
         Height          =   375
         Left            =   13440
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   240
         Width           =   2715
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   12
         Left            =   7800
         TabIndex        =   66
         Top             =   1200
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
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
         ButtonImage     =   "FrmUnitRent.frx":0EBE
         DrawFocusRectangle=   0   'False
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Height          =   315
         Left            =   11400
         TabIndex        =   68
         Top             =   600
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VSFlex8Ctl.VSFlexGrid fg 
         Height          =   1380
         Left            =   0
         TabIndex        =   76
         Top             =   120
         Width           =   7755
         _cx             =   13679
         _cy             =   2434
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
         BackColorFixed  =   -2147483633
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmUnitRent.frx":1458
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·«À»« "
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   3
         Left            =   10200
         TabIndex        =   72
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„—«›ﬁ"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   2
         Left            =   13560
         TabIndex        =   70
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·’›Â"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   1
         Left            =   14160
         TabIndex        =   69
         Top             =   600
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   2040
      Width           =   15135
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   11
         Left            =   120
         TabIndex        =   59
         Top             =   3120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
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
         ButtonImage     =   "FrmUnitRent.frx":1590
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ·Ì›Ê‰"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   45
         Left            =   7320
         TabIndex        =   64
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·«ﬁ«„…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   43
         Left            =   10320
         TabIndex        =   62
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·ﬂ›Ì·"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   44
         Left            =   14160
         TabIndex        =   60
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   1575
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   600
      Width           =   15135
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text27 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   12480
         Locked          =   -1  'True
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   120
         Width           =   1455
      End
      Begin XtremeSuiteControls.RadioButton RadioButton1 
         Height          =   255
         Left            =   3000
         TabIndex        =   52
         Top             =   1080
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "»—«"
         ForeColor       =   16711680
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄„·"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E2E9E9&
         Caption         =   "“Ì«—…"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   12480
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   960
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   3015
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   10
         Left            =   120
         TabIndex        =   31
         Top             =   3120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
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
         ButtonImage     =   "FrmUnitRent.frx":1B2A
         DrawFocusRectangle=   0   'False
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   315
         Left            =   12480
         TabIndex        =   41
         Top             =   600
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Dynamic_Byte.NourHijriCal Txt_DateExpEkamaH 
         Height          =   315
         Left            =   1560
         TabIndex        =   48
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   93519873
         CurrentDate     =   38784
      End
      Begin XtremeSuiteControls.RadioButton RadioButton2 
         Height          =   255
         Left            =   2160
         TabIndex        =   54
         Top             =   1080
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ÃÊ«"
         ForeColor       =   16711680
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RadioButton3 
         Height          =   255
         Left            =   1560
         TabIndex        =   55
         Top             =   1080
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "»Õ—«"
         ForeColor       =   16711680
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·”«ﬂ‰"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   24
         Left            =   6840
         TabIndex        =   124
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÃÂ… «·ﬁœÊ„"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   42
         Left            =   3720
         TabIndex        =   53
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”»» «·ﬁœÊ„"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   39
         Left            =   6840
         TabIndex        =   46
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»ÿ«ﬁ… «·«ÕÊ«·"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   38
         Left            =   7320
         TabIndex        =   45
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Â‰…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   27
         Left            =   10920
         TabIndex        =   40
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ·›Ê‰ «·„‰“·"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   37
         Left            =   13560
         TabIndex        =   39
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·„Ê»«Ì·"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   36
         Left            =   10920
         TabIndex        =   38
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " √—ÌŒ Ê„ﬂ«‰ ’œÊ—Â«"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   35
         Left            =   3480
         TabIndex        =   37
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ã‰”Ì…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   34
         Left            =   13560
         TabIndex        =   36
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„” √Ã—"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   33
         Left            =   10920
         TabIndex        =   35
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·⁄÷ÊÌ…"
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   31
         Left            =   13560
         TabIndex        =   33
         Top             =   240
         Width           =   1485
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   0
      Width           =   15135
      Begin VB.TextBox XPTxtID 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   12480
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   21
         Left            =   120
         TabIndex        =   26
         Top             =   3120
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   476
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
         ButtonImage     =   "FrmUnitRent.frx":20C4
         DrawFocusRectangle=   0   'False
      End
      Begin MSComCtl2.DTPicker XPDtbBill 
         Height          =   345
         Left            =   9840
         TabIndex        =   120
         Top             =   240
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   609
         _Version        =   393216
         Format          =   93519873
         CurrentDate     =   38784
      End
      Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
         Height          =   315
         Left            =   7920
         TabIndex        =   122
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ…"
         Height          =   285
         Index           =   23
         Left            =   10080
         TabIndex        =   121
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·⁄ﬁœ"
         Height          =   285
         Index           =   4
         Left            =   12870
         TabIndex        =   28
         Top             =   255
         Width           =   2205
      End
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   16200
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   15720
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   16200
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   16320
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   15420
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   2790
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6900
      Width           =   8745
      _cx             =   15425
      _cy             =   953
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   7230
         TabIndex        =   2
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   1
         Left            =   6375
         TabIndex        =   3
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   2
         Left            =   5535
         TabIndex        =   4
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   3
         Left            =   4680
         TabIndex        =   5
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   4
         Left            =   3825
         TabIndex        =   6
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   7
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   375
         Left            =   855
         TabIndex        =   8
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "„”«⁄œ…"
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   5
         Left            =   2760
         TabIndex        =   17
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   9
         Left            =   1920
         TabIndex        =   24
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   8580
      TabIndex        =   9
      Top             =   6480
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   16200
      TabIndex        =   10
      Top             =   3570
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   15840
      TabIndex        =   19
      Top             =   1920
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”«⁄…"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   20
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ⁄ﬁœ «· √ÃÌ— «· ”·”·Ì"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   32
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   1485
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ⁄ﬁœ «· √ÃÌ— «· ”·”·Ì"
      Height          =   285
      Index           =   29
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   1485
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   15090
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   11325
      TabIndex        =   16
      Top             =   6555
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2430
      TabIndex        =   15
      Top             =   6630
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   690
      TabIndex        =   14
      Top             =   6630
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   13
      Top             =   6660
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   12
      Top             =   6660
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   16350
      TabIndex        =   11
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmUnitRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Form_Load()
Resize_Form Me
End Sub
