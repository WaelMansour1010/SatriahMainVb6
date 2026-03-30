VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form Projects 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7710
   ClientLeft      =   3525
   ClientTop       =   1470
   ClientWidth     =   22320
   Icon            =   "project2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   22320
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "‰Ê⁄ «·„‘—Ê⁄"
      ClipControls    =   0   'False
      Height          =   615
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   222
      Top             =   0
      Width           =   3255
      Begin VB.OptionButton Ptype 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«›  «ÕÌ"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   224
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Ptype 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ÃœÌœ"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   223
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "«·„⁄œ«  Ê «·√·« "
      Height          =   3615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   210
      Top             =   -4680
      Visible         =   0   'False
      Width           =   22215
      Begin VB.TextBox Text14 
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
         Height          =   300
         Left            =   2640
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VB.TextBox Text11 
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
         Height          =   300
         Left            =   4200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   216
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VB.TextBox Text9 
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
         Height          =   300
         Left            =   5760
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   215
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VB.TextBox Text8 
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
         Height          =   300
         Left            =   7320
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   211
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid4 
         Height          =   2340
         Left            =   120
         TabIndex        =   212
         Top             =   360
         Width           =   22080
         _cx             =   38947
         _cy             =   4128
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
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"project2.frx":000C
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
      Begin ALLButtonS.ALLButton opr_Expenses 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   213
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "—ÃÊ⁄ ··⁄„·Ì« "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":01B3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì« "
         Height          =   255
         Index           =   7
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   214
         Top             =   2760
         Width           =   2535
      End
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "›⁄·Ì"
      Height          =   255
      Left            =   9120
      TabIndex        =   208
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00C0FFFF&
      Caption         =   " ﬁœÌ—Ì"
      Height          =   255
      Left            =   10320
      TabIndex        =   207
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·œ›⁄« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3645
      Index           =   2
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   152
      Top             =   3360
      Visible         =   0   'False
      Width           =   22215
      Begin VB.TextBox Text12 
         Height          =   735
         Left            =   960
         TabIndex        =   153
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VSFlex8Ctl.VSFlexGrid GridSub 
         Height          =   2835
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   21315
         _cx             =   37597
         _cy             =   5001
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
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   25
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"project2.frx":01CF
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Left            =   21480
         TabIndex        =   225
         Top             =   720
         Width           =   690
         _ExtentX        =   1217
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
         ButtonImage     =   "project2.frx":05A8
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   21840
         TabIndex        =   154
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  „Õ«”»Ì…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3645
      Index           =   3
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   155
      Top             =   3360
      Visible         =   0   'False
      Width           =   22215
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2925
         Index           =   7
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   158
         Top             =   480
         Width           =   15375
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—’Ìœ «·√›  «ÕÏ  ··«ÃÊ—"
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
            Height          =   1335
            Index           =   1
            Left            =   12240
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   1440
            Width           =   3075
            Begin VB.OptionButton OptType3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "€Ì— „Õœœ"
               Height          =   255
               Index           =   2
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   196
               Top             =   210
               Width           =   915
            End
            Begin VB.OptionButton OptType3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ«∆‰"
               Height          =   255
               Index           =   1
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   195
               Top             =   210
               Width           =   915
            End
            Begin VB.OptionButton OptType3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÌ‰"
               Height          =   255
               Index           =   0
               Left            =   1950
               RightToLeft     =   -1  'True
               TabIndex        =   194
               Top             =   210
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.TextBox TxtOpenBalance3 
               Alignment       =   2  'Center
               Height          =   345
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   193
               Top             =   480
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker Dtp3 
               Height          =   330
               Left            =   150
               TabIndex        =   197
               Top             =   900
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               CalendarBackColor=   12648447
               CustomFormat    =   "yyyy/M/d"
               Format          =   100597763
               CurrentDate     =   38718
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «· ”ÃÌ·"
               Height          =   315
               Index           =   11
               Left            =   1770
               RightToLeft     =   -1  'True
               TabIndex        =   199
               Top             =   960
               Width           =   1125
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—’Ìœ "
               Height          =   345
               Index           =   10
               Left            =   1770
               RightToLeft     =   -1  'True
               TabIndex        =   198
               Top             =   510
               Width           =   1125
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—’Ìœ «·√›  «ÕÏ „” Õ·’« "
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
            Height          =   1335
            Index           =   0
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   1440
            Width           =   3075
            Begin VB.OptionButton OptType4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "€Ì— „Õœœ"
               Height          =   255
               Index           =   2
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   188
               Top             =   210
               Width           =   915
            End
            Begin VB.OptionButton OptType4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ«∆‰"
               Height          =   255
               Index           =   1
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Top             =   210
               Width           =   915
            End
            Begin VB.OptionButton OptType4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÌ‰"
               Height          =   255
               Index           =   0
               Left            =   1950
               RightToLeft     =   -1  'True
               TabIndex        =   186
               Top             =   210
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.TextBox TxtOpenBalance4 
               Alignment       =   2  'Center
               Height          =   345
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   185
               Top             =   480
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker Dtp4 
               Height          =   330
               Left            =   150
               TabIndex        =   189
               Top             =   900
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               CalendarBackColor=   12648447
               CustomFormat    =   "yyyy/M/d"
               Format          =   100597763
               CurrentDate     =   38718
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «· ”ÃÌ·"
               Height          =   315
               Index           =   9
               Left            =   1770
               RightToLeft     =   -1  'True
               TabIndex        =   191
               Top             =   900
               Width           =   1125
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—’Ìœ "
               Height          =   345
               Index           =   8
               Left            =   1770
               RightToLeft     =   -1  'True
               TabIndex        =   190
               Top             =   510
               Width           =   1125
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—’Ìœ «·√›  «ÕÏ „’—Ê›« "
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
            Height          =   1335
            Index           =   8
            Left            =   12300
            RightToLeft     =   -1  'True
            TabIndex        =   176
            Top             =   120
            Width           =   3075
            Begin VB.OptionButton OptType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "€Ì— „Õœœ"
               Height          =   255
               Index           =   2
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   180
               Top             =   210
               Width           =   915
            End
            Begin VB.OptionButton OptType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ«∆‰"
               Height          =   255
               Index           =   1
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   179
               Top             =   210
               Width           =   915
            End
            Begin VB.OptionButton OptType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÌ‰"
               Height          =   255
               Index           =   0
               Left            =   1950
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Top             =   210
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.TextBox TxtOpenBalance 
               Alignment       =   2  'Center
               Height          =   345
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   480
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker Dtp 
               Height          =   330
               Left            =   150
               TabIndex        =   181
               Top             =   900
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               CalendarBackColor=   12648447
               CustomFormat    =   "yyyy/M/d"
               Format          =   100597763
               CurrentDate     =   38718
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «· ”ÃÌ·"
               Height          =   315
               Index           =   13
               Left            =   1770
               RightToLeft     =   -1  'True
               TabIndex        =   183
               Top             =   900
               Width           =   1125
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—’Ìœ "
               Height          =   345
               Index           =   14
               Left            =   1770
               RightToLeft     =   -1  'True
               TabIndex        =   182
               Top             =   510
               Width           =   1125
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—’Ìœ «·√›  «ÕÏ «Ì—«œ« "
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
            Height          =   1335
            Index           =   9
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   168
            Top             =   120
            Width           =   3075
            Begin VB.TextBox TxtOpenBalance1 
               Alignment       =   2  'Center
               Height          =   345
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   172
               Top             =   480
               Width           =   1575
            End
            Begin VB.OptionButton OptType1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÌ‰"
               Height          =   255
               Index           =   0
               Left            =   1950
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   210
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.OptionButton OptType1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ«∆‰"
               Height          =   255
               Index           =   1
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   170
               Top             =   210
               Width           =   915
            End
            Begin VB.OptionButton OptType1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "€Ì— „Õœœ"
               Height          =   255
               Index           =   2
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   169
               Top             =   210
               Width           =   915
            End
            Begin MSComCtl2.DTPicker Dtp1 
               Height          =   330
               Left            =   150
               TabIndex        =   173
               Top             =   900
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               CalendarBackColor=   12648447
               CustomFormat    =   "yyyy/M/d"
               Format          =   100597763
               CurrentDate     =   38718
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—’Ìœ "
               Height          =   345
               Index           =   15
               Left            =   1770
               RightToLeft     =   -1  'True
               TabIndex        =   175
               Top             =   510
               Width           =   1125
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «· ”ÃÌ·"
               Height          =   315
               Index           =   16
               Left            =   1770
               RightToLeft     =   -1  'True
               TabIndex        =   174
               Top             =   900
               Width           =   1125
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—’Ìœ «·√›  «ÕÏ „Ê«œ"
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
            Height          =   1335
            Index           =   10
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Top             =   120
            Width           =   3075
            Begin VB.OptionButton OptType2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "€Ì— „Õœœ"
               Height          =   255
               Index           =   2
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   210
               Width           =   915
            End
            Begin VB.OptionButton OptType2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ«∆‰"
               Height          =   255
               Index           =   1
               Left            =   990
               RightToLeft     =   -1  'True
               TabIndex        =   163
               Top             =   210
               Width           =   915
            End
            Begin VB.OptionButton OptType2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÌ‰"
               Height          =   255
               Index           =   0
               Left            =   1920
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   210
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.TextBox TxtOpenBalance2 
               Alignment       =   2  'Center
               Height          =   345
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   480
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker Dtp2 
               Height          =   330
               Left            =   150
               TabIndex        =   165
               Top             =   900
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               Enabled         =   0   'False
               CalendarBackColor=   12648447
               CustomFormat    =   "yyyy/M/d"
               Format          =   100597763
               CurrentDate     =   38718
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «· ”ÃÌ·"
               Height          =   315
               Index           =   17
               Left            =   1770
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   900
               Width           =   1125
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—’Ìœ "
               Height          =   345
               Index           =   18
               Left            =   1770
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   510
               Width           =   1125
            End
         End
         Begin VB.TextBox txtopening_balance_voucher_id 
            Height          =   735
            Left            =   960
            TabIndex        =   159
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   14760
            TabIndex        =   200
            Top             =   -1080
            Width           =   375
         End
      End
      Begin VB.TextBox Text7 
         Height          =   735
         Left            =   7200
         TabIndex        =   156
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   14880
         TabIndex        =   157
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   360
      TabIndex        =   130
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame13 
      Height          =   2775
      Left            =   0
      TabIndex        =   92
      Top             =   600
      Width           =   22215
      Begin VB.TextBox txt_total_discount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   14040
         TabIndex        =   233
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox TXTprojectnamee 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   14040
         TabIndex        =   232
         Top             =   840
         Width           =   5895
      End
      Begin VB.TextBox txtid 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   17640
         TabIndex        =   230
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox TxtProjectCosts 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   17640
         TabIndex        =   229
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox total_after_discount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   17640
         Locked          =   -1  'True
         TabIndex        =   228
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   17640
         TabIndex        =   227
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox TxtDiscountPercentage 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   14040
         TabIndex        =   9
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox TxtCustCode2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   11760
         TabIndex        =   205
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   11760
         TabIndex        =   203
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox TxtCustCode1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   11760
         TabIndex        =   202
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtCustCode 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   11760
         TabIndex        =   201
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5760
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2280
         Width           =   7095
      End
      Begin MSDataListLib.DataCombo DcAccount1 
         Height          =   315
         Left            =   22320
         TabIndex        =   116
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   97
         Text            =   "Text2"
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TXTprojectname 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   14040
         TabIndex        =   5
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox txt_project_id 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   22200
         TabIndex        =   96
         Top             =   -120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "»‰Êœ"
         Height          =   195
         Left            =   11640
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   3120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "«’‰«›"
         Height          =   195
         Left            =   10560
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   3120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄„«·Â"
         Height          =   195
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   3120
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DcAccount2 
         Height          =   315
         Left            =   5760
         TabIndex        =   6
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcAccount3 
         Height          =   315
         Left            =   22200
         TabIndex        =   117
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcAccount4 
         Height          =   315
         Left            =   5760
         TabIndex        =   7
         Top             =   840
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCPreFix 
         Height          =   315
         Left            =   16560
         TabIndex        =   0
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcCurrency 
         Height          =   315
         Left            =   5760
         TabIndex        =   4
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Height          =   315
         Left            =   9600
         TabIndex        =   2
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DataCombo5 
         Height          =   315
         Left            =   7560
         TabIndex        =   3
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   14040
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTEnddate 
         Height          =   285
         Left            =   14040
         TabIndex        =   11
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         Format          =   100597761
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   9840
         TabIndex        =   137
         Top             =   2760
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   100597761
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcEmp 
         Height          =   315
         Left            =   5760
         TabIndex        =   10
         Top             =   1560
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   285
         Left            =   14040
         TabIndex        =   12
         Top             =   2280
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         Format          =   100597761
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcEmp1 
         Height          =   315
         Left            =   5760
         TabIndex        =   8
         Top             =   1200
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbDept 
         Height          =   315
         Left            =   5760
         TabIndex        =   226
         Top             =   1920
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DTStartDate 
         Height          =   285
         Left            =   17640
         TabIndex        =   231
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         Format          =   100597761
         CurrentDate     =   38784
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«œ«—Â"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12720
         TabIndex        =   54
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "‰”»…«·Œ’„"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   16320
         TabIndex        =   206
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„‰œÊ»"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12840
         TabIndex        =   204
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·„‘—Ê⁄ «‰Ã·Ì“Ì"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   20160
         TabIndex        =   149
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "√ﬁ—» ‰Â«Ì…"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   16680
         TabIndex        =   144
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "„œÌ— «·„Êﬁ⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12840
         TabIndex        =   143
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "„·«ÕŸ« "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12720
         TabIndex        =   140
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "„œ… «·„‘—Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   20160
         TabIndex        =   139
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   " «ﬁ—» ‰Â«Ì…"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   11040
         TabIndex        =   138
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«‰ Â«¡ "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   16800
         TabIndex        =   136
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   " «—ÌŒ «·»œ«Ì…"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   20160
         TabIndex        =   135
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Õ«·… «·„‘—Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   15360
         TabIndex        =   113
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂÊœ «·„‘—Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   20160
         TabIndex        =   112
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·„‘—Ê⁄ ⁄—»Ì"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   20160
         TabIndex        =   111
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "‰Ê⁄ «·⁄ﬁœ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   110
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·⁄„Ì·"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   21960
         TabIndex        =   109
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄„Ì· «·‰Â«∆Ì"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12840
         TabIndex        =   108
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÌ„… «·„‘—Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   20160
         TabIndex        =   107
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄„Ì· «·»«ÿ‰"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12840
         TabIndex        =   106
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·⁄„Ì·"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   21960
         TabIndex        =   105
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "«·›—⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12840
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÌ„… «·Œ’„"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   16320
         TabIndex        =   103
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÌ„… «·„‘—Ê⁄  »⁄œ «·Œ’„"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   20160
         TabIndex        =   102
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄„·Â"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   $"project2.frx":0B42
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
         Height          =   1980
         Index           =   5
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   4
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   120
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   2055
         Left            =   120
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "„”·”· «·„‘—Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   22200
         TabIndex        =   98
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "«·„’—Ê›« "
      Height          =   3615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   87
      Top             =   3240
      Visible         =   0   'False
      Width           =   22215
      Begin VB.TextBox txt_expenses_total 
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
         Height          =   300
         Left            =   6960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
         Height          =   2340
         Left            =   120
         TabIndex        =   88
         Top             =   360
         Width           =   21960
         _cx             =   38735
         _cy             =   4128
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
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"project2.frx":0C57
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
      Begin ALLButtonS.ALLButton opr_Expenses 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   91
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "—ÃÊ⁄ ··⁄„·Ì« "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":0D88
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì ﬁÌ„… «·„’—Ê›« "
         Height          =   255
         Index           =   6
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   2760
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "„Ê«œ «·⁄„·Ì… —ﬁ„"
      Height          =   3615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3240
      Visible         =   0   'False
      Width           =   22215
      Begin VB.TextBox XPTxtDiscountVal 
         Height          =   375
         Left            =   11400
         TabIndex        =   147
         Text            =   "Text7"
         Top             =   3240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox XPCboDiscountType 
         Height          =   315
         Left            =   10560
         TabIndex        =   146
         Text            =   "Combo2"
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   480
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   1320
         Width           =   1530
      End
      Begin VB.TextBox XPTxtBillID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   360
         Left            =   720
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox TxtFillData 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox XPTxtSum 
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
         Height          =   300
         Left            =   480
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1530
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   480
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   840
         Width           =   1530
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   480
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   480
         Width           =   1530
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   690
         Index           =   2
         Left            =   3600
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   18555
         _cx             =   32729
         _cy             =   1217
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
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   8385
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   300
            Width           =   2265
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5895
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   218
            Top             =   300
            Width           =   2235
         End
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   945
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   300
            Width           =   2235
         End
         Begin VB.TextBox TxtSerial 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Left            =   5505
            MaxLength       =   20
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   -300
            Width           =   3195
         End
         Begin VB.TextBox TxtQuantity 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   3255
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   300
            Width           =   2250
         End
         Begin VB.ComboBox CboItemCase 
            Height          =   315
            Left            =   8775
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   -420
            Width           =   2400
         End
         Begin MSDataListLib.DataCombo DCboItemsName 
            Height          =   315
            Left            =   11190
            TabIndex        =   39
            Top             =   300
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboItemsCode 
            Height          =   315
            Left            =   14700
            TabIndex        =   40
            Top             =   300
            Width           =   3720
            _ExtentX        =   6562
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdAdd 
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   270
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            ButtonImage     =   "project2.frx":0DA4
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
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ„Ì…  ﬁœÌ—Ì"
            Height          =   255
            Index           =   19
            Left            =   8985
            RightToLeft     =   -1  'True
            TabIndex        =   221
            Top             =   0
            Width           =   1590
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄—  ﬁœÌ—Ì"
            Height          =   255
            Index           =   12
            Left            =   6585
            RightToLeft     =   -1  'True
            TabIndex        =   220
            Top             =   0
            Width           =   1740
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄— «·›⁄·Ì"
            Height          =   255
            Index           =   26
            Left            =   1455
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   0
            Width           =   1725
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ„Ì… «·›⁄·Ì…"
            Height          =   255
            Index           =   27
            Left            =   3855
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ì—Ì«·"
            Height          =   255
            Index           =   28
            Left            =   6015
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   -600
            Width           =   2715
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·’‰›"
            Height          =   255
            Index           =   29
            Left            =   8985
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   -720
            Width           =   2205
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈”„ «·’‰›"
            Height          =   255
            Index           =   30
            Left            =   11550
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   0
            Width           =   3150
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   31
            Left            =   15300
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   0
            Width           =   3225
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid FG 
         Height          =   2145
         Left            =   3600
         TabIndex        =   48
         Top             =   960
         Width           =   18555
         _cx             =   32729
         _cy             =   3784
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
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"project2.frx":113E
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
         WallPaperAlignment=   0
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin ALLButtonS.ALLButton opr_items 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   79
         Top             =   2640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "—ÃÊ⁄ ··⁄„·Ì« "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":138D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblTotalQty 
         Caption         =   "Label38"
         Height          =   135
         Left            =   12120
         TabIndex        =   145
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ÿ·Ê»"
         Height          =   255
         Index           =   3
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label LblItemsCount 
         Caption         =   "Label27"
         Height          =   135
         Left            =   240
         TabIndex        =   64
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì ﬁÌ„… «·«’‰«›"
         Height          =   255
         Index           =   2
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ÕÃÊ“"
         Height          =   255
         Index           =   1
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ Ê›—"
         Height          =   255
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "«”„«¡ «·⁄«„·Ì‰ ›Ì «·„‘—Ê⁄"
      Height          =   3615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   3360
      Visible         =   0   'False
      Width           =   22215
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   118
         Top             =   240
         Width           =   21975
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            Caption         =   " ﬁœÌ—Ì"
            Height          =   255
            Left            =   13800
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   120
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option5 
            Alignment       =   1  'Right Justify
            Caption         =   " Œ’Ì’ ›⁄·Ì"
            Height          =   255
            Left            =   12120
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox TxtCount 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6960
            TabIndex        =   121
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "«œ—«Ã"
            Height          =   255
            Left            =   960
            TabIndex        =   120
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox TxtEmpcount 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   9480
            TabIndex        =   119
            Top             =   120
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo dcJobTypeName 
            Height          =   315
            Left            =   15720
            TabIndex        =   124
            Top             =   120
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   4200
            TabIndex        =   125
            Top             =   120
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   100597761
            CurrentDate     =   38784
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "«Œ — «·„Â‰… «·„ÿ·Ê»…"
            Height          =   255
            Left            =   20160
            TabIndex        =   129
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄œœ «·⁄„«·"
            Height          =   255
            Left            =   11040
            TabIndex        =   128
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   " «—ÌŒ «· Œ’Ì’"
            Height          =   255
            Left            =   5640
            TabIndex        =   127
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄œœ «·«Ì«„"
            Height          =   255
            Left            =   8520
            TabIndex        =   126
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.TextBox txt_emp_salary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   6360
         TabIndex        =   86
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txt_employee_count 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   9960
         TabIndex        =   84
         Top             =   3000
         Width           =   1575
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   1860
         Left            =   120
         TabIndex        =   71
         Top             =   960
         Width           =   22080
         _cx             =   38947
         _cy             =   3281
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   0   'False
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
         Rows            =   3
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"project2.frx":13A9
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
      Begin ALLButtonS.ALLButton opr_emplyees_name 
         Height          =   375
         Left            =   120
         TabIndex        =   82
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "—ÃÊ⁄ ··⁄„·Ì« "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":153D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label29 
         Caption         =   "ﬁÌ„… «ÃÊ— «·⁄„«·"
         Height          =   255
         Left            =   8040
         TabIndex        =   85
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "«Ã„«·Ì ⁄œœ «·⁄„·"
         Height          =   255
         Left            =   11640
         TabIndex        =   83
         Top             =   3120
         Width           =   1815
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "⁄„·Ì«  ﬂ· »‰œ"
      Height          =   3615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   3360
      Visible         =   0   'False
      Width           =   22215
      Begin VB.Frame Frame4 
         Caption         =   "œ·«·«  «·«·Ê«‰"
         Height          =   615
         Left            =   12360
         RightToLeft     =   -1  'True
         TabIndex        =   141
         Top             =   3000
         Width           =   2895
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Caption         =   "Õ—Ã"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox TXTNoOFWeek 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   132
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txt_opr_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   2760
         Width           =   3015
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
         Height          =   2340
         Left            =   0
         TabIndex        =   75
         Top             =   240
         Width           =   22080
         _cx             =   38947
         _cy             =   4128
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
         Rows            =   3
         Cols            =   40
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"project2.frx":1559
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   -1  'True
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
      Begin ALLButtonS.ALLButton terms_operations 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   77
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "—ÃÊ⁄ ··»‰Êœ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1B13
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton opr_items 
         Height          =   255
         Index           =   0
         Left            =   20280
         TabIndex        =   78
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "„Ê«œ "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1B2F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton employee_details 
         Height          =   255
         Left            =   16440
         TabIndex        =   80
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "»Ì«‰«  «·⁄„«·…"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1B4B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton opr_Expenses 
         Height          =   255
         Index           =   0
         Left            =   14520
         TabIndex        =   81
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "„’«—Ì›"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1B67
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton CMDViewGantt 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   133
         Top             =   3120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "⁄—÷ «·Ã«‰ "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1B83
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton opr_Expenses 
         Height          =   255
         Index           =   2
         Left            =   18360
         TabIndex        =   209
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "«·„⁄œ« "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1B9F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«”»Ê⁄"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   134
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì «·„œ…"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5760
         TabIndex        =   131
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì «· ﬂ·›…"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5760
         TabIndex        =   76
         Top             =   2760
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "«·»‰Êœ"
      Height          =   3615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   3360
      Width           =   22215
      Begin VB.TextBox txt_sub_discount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txt_total_sum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2280
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         TabIndex        =   61
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txt_sub_net 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   2760
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
         Height          =   2340
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   22065
         _cx             =   38920
         _cy             =   4128
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
         Rows            =   3
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"project2.frx":1BBB
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
         Begin VB.PictureBox PicDes 
            BorderStyle     =   0  'None
            Height          =   1635
            Left            =   240
            RightToLeft     =   -1  'True
            ScaleHeight     =   1635
            ScaleWidth      =   10485
            TabIndex        =   51
            Top             =   960
            Visible         =   0   'False
            Width           =   10485
            Begin VB.TextBox TxtDes 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   1125
               Left            =   30
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   52
               Top             =   360
               Visible         =   0   'False
               Width           =   2115
            End
            Begin VB.Label LblDes 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000C&
               Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
               ForeColor       =   &H0000C8FF&
               Height          =   315
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   0
               Width           =   4485
            End
         End
      End
      Begin ALLButtonS.ALLButton terms_operations 
         Height          =   375
         Index           =   0
         Left            =   19920
         TabIndex        =   72
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "⁄„·Ì«  «·»‰œ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1E1E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "«·«Ã„«·Ì"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3360
         TabIndex        =   62
         Top             =   2760
         Width           =   975
      End
   End
   Begin VB.TextBox TxtModFlg 
      Height          =   285
      Left            =   3480
      TabIndex        =   31
      Text            =   "txtmodflag"
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   1920
      TabIndex        =   19
      Top             =   7080
      Width           =   13215
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   1
         Left            =   8160
         TabIndex        =   20
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Õ›Ÿ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1E3A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   21
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "«·„—›ﬁ« "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1E56
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   0
         Left            =   11760
         TabIndex        =   22
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "ÃœÌœ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8454016
         BCOLO           =   8454016
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1E72
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   24
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "»ÕÀ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1E8E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command3 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "«· ﬁ—Ì—"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1EAA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   -1  'True
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   3
         Left            =   10560
         TabIndex        =   32
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   " ⁄œÌ·"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1EC6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   6
         Left            =   9360
         TabIndex        =   114
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   " —«Ã⁄"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1EE2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   7
         Left            =   7080
         TabIndex        =   115
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Õ–›"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1EFE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   4
         Left            =   3960
         TabIndex        =   150
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "«·«—’œ… «·«›  «ÕÌ…"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1F1A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   375
         Index           =   8
         Left            =   2520
         TabIndex        =   151
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "»Ì«‰«  «·œ›⁄« "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "project2.frx":1F36
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Caption         =   "Label2"
         Height          =   15
         Left            =   -120
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
   End
   Begin ALLButtonS.ALLButton CMD_language 
      Height          =   495
      Left            =   0
      TabIndex        =   18
      ToolTipText     =   "Language  «··€…"
      Top             =   480
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "EN"
      ENAB            =   -1  'True
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
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "project2.frx":1F52
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   -720
      TabIndex        =   15
      Top             =   6360
      Width           =   1935
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   0
      Left            =   1665
      TabIndex        =   25
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
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
      ButtonImage     =   "project2.frx":1F6E
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
      Height          =   345
      Index           =   2
      Left            =   600
      TabIndex        =   26
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
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
      ButtonImage     =   "project2.frx":2308
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
      Height          =   345
      Index           =   1
      Left            =   2190
      TabIndex        =   27
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
      ButtonStyle     =   1
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
      ButtonImage     =   "project2.frx":26A2
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      Alignment       =   0
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      RightToLeft     =   -1  'True
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   3
      Left            =   1125
      TabIndex        =   28
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
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
      ButtonImage     =   "project2.frx":2A3C
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   315
      Left            =   0
      TabIndex        =   148
      Top             =   7080
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   100597761
      CurrentDate     =   38784
   End
   Begin VB.Label Label25 
      Caption         =   "Label25"
      Height          =   375
      Left            =   1080
      TabIndex        =   29
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "     »Ì«‰«  «·„‘«—Ì⁄      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   22215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   11880
      TabIndex        =   16
      Top             =   9840
      Width           =   855
   End
End
Attribute VB_Name = "Projects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As String
Dim last_root As Integer
Dim last_geeral As Integer
Dim last_branch As Integer
Dim mod_flad As String
Dim first_run  As Boolean
Dim fullcode As String
Dim test1 As Boolean
Dim rs As ADODB.Recordset
Dim RsDev As ADODB.Recordset
Dim RsDevsub As ADODB.Recordset
Dim NewGrid As New ClsGrid
Dim RSTransDetails As ADODB.Recordset
Dim RsDetails As ADODB.Recordset
Dim current_terms As String
Public ProjectDes_ID As Integer
Dim current_opr As String
Public LngRow As Double
Dim Account_Code_dynamic1 As String
Dim Account_Code_dynamic2 As String
Dim Account_Code_dynamic3 As String
Dim Account_Code_dynamic4 As String
Dim Account_Code_dynamic5 As String
Dim Account_Code_dynamic1C As String
Dim Account_Code_dynamic2C As String
Dim Account_Code_dynamic3C As String
Dim Account_Code_dynamic4C As String
Dim Account_Code_dynamic5C As String

Private Sub SaveData()

    If SystemOptions.UserInterface = EnglishInterface Then

        If DcAccount1.BoundText = "" Then MsgBox "Must Specify Client Name", vbCritical: Exit Sub
        If DcCurrency.BoundText = "" Then
            MsgBox "Must Specify  Currency Name", vbCritical
            DcCurrency.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If DcAccount2.BoundText = "" Then MsgBox "Must Specify client Name", vbCritical: Exit Sub
        If TXTprojectname.text = "" Then MsgBox "Must Specify project Name", vbCritical: Exit Sub

        If Not IsNumeric(TxtProjectCosts.text) Then MsgBox "Must Specify project cost", vbCritical: Exit Sub
  
        If Not IsNumeric(txt_total_discount.text) Then MsgBox "discount must be numeric", vbCritical:  txt_total_discount.text = 0: Exit Sub
        If Not IsNumeric(Dcbranch.BoundText) Then MsgBox "Must select Branch", vbCritical: Exit Sub
        If DataCombo1.text = "" Then MsgBox "Must Specify project Status", vbCritical: Exit Sub
    Else

   '     If DCAccount1.BoundText = "" Then MsgBox "·«»œ „‰  ÕœÌœ «”„ „ﬁ«Ê· «·»«ÿ‰  ", vbCritical: Exit Sub
        If DcAccount2.BoundText = "" Then MsgBox "·«»œ „‰  ÕœÌœ «”„ «·⁄„Ì· «·‰Â«∆Ì", vbCritical: Exit Sub
        If TXTprojectname.text = "" Then MsgBox "·«»œ „‰  ÕœÌœ «”„   «·„‘—Ê⁄", vbCritical: Exit Sub

        If Not IsNumeric(TxtProjectCosts.text) Then MsgBox "ÌÃ»  ÕœÌœ ﬁÌ„… «·„‘—Ê⁄", vbCritical: Exit Sub
  
        If Not IsNumeric(txt_total_discount.text) Then MsgBox "·«»œ „‰  ÕœÌœ «·Œ’„", vbCritical:  txt_total_discount.text = 0: Exit Sub
        If DcCurrency.BoundText = "" Then
            MsgBox "Õœœ «·⁄„·… «Ê·«", vbCritical
            DcCurrency.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If Not IsNumeric(Dcbranch.BoundText) Then MsgBox "Õœœ «·›—⁄ «Ê·«", vbCritical: Exit Sub
        If DataCombo1.text = "" Then MsgBox "Õœœ Õ«·… «·„‘—Ê⁄", vbCritical: Exit Sub

    End If
  
    'If txtid.text = "" Then
    'txtid.text = get_code
    'End If

    Dim currentcode As String

    If txtid.text = "" Then
        currentcode = get_coding(branch_id, "projects", 0, Me.DCPreFix.text)

        If currentcode = "miniError" Then
            MsgBox "⁄œœ «·Œ«‰«  «· Ì ﬁ„  » ÕœÌœ…  ·Â–« ««ﬂÊœ ’€Ì—… Ãœ« Ì—ÃÌ  €ÌÌ—Â« ›Ì ‘«‘…  ﬂÊÌœ «·ÕﬁÊ· «Ê «·« ’«· »„”∆Ê· «·‰Ÿ«„"
            Exit Sub
                        
        ElseIf currentcode = "Manual" Then
            MsgBox "«œŒ· «·ﬂÊœ ÌœÊÌ« ﬂ„« Õœœ  ›Ì  ﬂÊÌœ «·”‰œ« "
            Exit Sub
        Else
            txtid = currentcode
        End If
    End If

    If txtid.text = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            MsgBox "Must enter project code or define coding in your System", vbCritical: Exit Sub
        Else
            MsgBox "·«»œ „‰ ﬂ «»… —ﬁ„ ··„‘—Ê⁄ ·«‰ﬂ ·„  Õœœ  ﬂÊÌœ «·Ì ·…", vbCritical: Exit Sub
        End If
    End If
  
    If Me.OptType(2).value = False Then
        If val(Me.TxtOpenBalance.text) = 0 Then
            Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ  ··„’—Ê›«   ...!!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If TxtOpenBalance.Enabled = True Then
                TxtOpenBalance.SetFocus
            End If

            Exit Sub
        End If
    End If
    
    If Me.OptType1(2).value = False Then
        If val(Me.TxtOpenBalance1.text) = 0 Then
            Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··«ÃÊ— ··«Ì—«œ«   ...!!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If TxtOpenBalance1.Enabled = True Then
                TxtOpenBalance1.SetFocus
            End If

            Exit Sub
        End If
    End If
            
    If Me.OptType2(2).value = False Then
        If val(Me.TxtOpenBalance2.text) = 0 Then
            Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··„Ê«œ ...!!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If TxtOpenBalance2.Enabled = True Then
                TxtOpenBalance2.SetFocus
            End If

            Exit Sub
        End If
    End If
     
    If Me.OptType3(2).value = False Then
        If val(Me.TxtOpenBalance3.text) = 0 Then
            Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··«ÕÊ— ...!!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If TxtOpenBalance3.Enabled = True Then
                TxtOpenBalance3.SetFocus
            End If

            Exit Sub
        End If
    End If
     
    If Me.OptType4(2).value = False Then
        If val(Me.TxtOpenBalance4.text) = 0 Then
            Msg = "ÌÃ» ﬂ «»Â ﬁÌ„… «·—’Ìœ ··„” Œ·’«  ...!!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

            If TxtOpenBalance4.Enabled = True Then
                TxtOpenBalance4.SetFocus
            End If

            Exit Sub
        End If
    End If
     
 '   total_after_discount = val(TxtProjectCosts) - val(txt_total_discount)
    'terms_operations_Click (1)
    'opr_items_Click (1)
    'opr_Expenses_Click (1)
    'opr_emplyees_name_Click

    If TxtModFlg.text = "N" Then

        'XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
        'Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=3"))
        If create_accounts = False Then Exit Sub
       
        rs.AddNew
        rs("id").value = val(Me.txt_project_id.text)
        rs("expanses_account").value = Account_Code_dynamic1C '  IIf(Trim$(Me.EXPANSES.text) = "", Null, Trim$(Me.EXPANSES.text))
        rs("REVENUE_account").value = Account_Code_dynamic2C ' IIf(Trim$(Me.REVENUE.text) = "", Null, Trim$(Me.REVENUE.text))
        rs("Material_account").value = Account_Code_dynamic3C '  IIf(Trim$(Me.Material.text) = "", Null, Trim$(Me.Material.text))
        rs("Salary_account").value = Account_Code_dynamic4C ' IIf(Trim$(Me.salary.text) = "", Null, Trim$(Me.salary.text))
        rs("legal").value = Account_Code_dynamic5C ' IIf(Trim$(Me.legal.text) = "", Null, Trim$(Me.legal.text))
                    
    Else

        If Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From projects_des Where project_id =" & val(Me.txt_project_id.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                    
           StrSQL = "Delete From Projectssub Where projectid =" & val(Me.txt_project_id.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
            Cn.Execute StrSQL, , adExecuteNoRecords

            If Not IsNull(rs("expanses_account").value) And rs("expanses_account").value <> "" Then
                    
                ModAccounts.EditAccount rs("expanses_account").value, TXTprojectname & " - „’—Ê›«  ", TXTprojectname & "- Expenses", , , , , , , , , , , , , , , , , True
            End If
                         
            If Not IsNull(rs("REVENUE_account").value) And rs("REVENUE_account").value <> "" Then
                    
                ModAccounts.EditAccount rs("REVENUE_account").value, TXTprojectname & " - «Ì—«œ«  ", TXTprojectname & "- Revenue", , , , , , , , , , , , , , , , , True
            End If
                         
            If Not IsNull(rs("Material_account").value) And rs("Material_account").value <> "" Then
                    
                ModAccounts.EditAccount rs("Material_account").value, TXTprojectname & " - „Ê«œ ", TXTprojectname & "- Material", , , , , , , , , , , , , , , , , True
            End If
                         
            If Not IsNull(rs("Salary_account").value) And rs("Salary_account").value <> "" Then
       
                ModAccounts.EditAccount rs("Salary_account").value, TXTprojectname & " - «ÃÊ— ", TXTprojectname & "- Salary", , , , , , , , , , , , , , , , , True
            End If
            
            If Not IsNull(rs("legal").value) And rs("legal").value <> "" Then
       
                ModAccounts.EditAccount rs("legal").value, TXTprojectname & " - „” Œ·’«  ", TXTprojectname & "- bill", , , , , , , , , , , , , , , , , True
            End If
                         
        End If
    End If
          rs("EmpId").value = IIf(Me.DcEmp.BoundText = "", Null, (Me.DcEmp.BoundText))

     rs("EmpId1").value = IIf(Me.DcEmp1.BoundText = "", Null, (Me.DcEmp1.BoundText))
    rs("StartDate").value = DTStartDate.value
     rs("Enddate").value = DTEnddate.value

    
    rs("End_user_Account").value = IIf(Trim(Me.DcAccount1.text) = "", Null, Trim(Me.DcAccount1.text))
    rs("End_user_name").value = IIf(DcAccount2.text = "", Null, DcAccount2.text)
    
        rs("End_user_id").value = IIf(DcAccount2.BoundText = "", Null, DcAccount2.BoundText)

        rs("sub_contractor_id").value = IIf(DcAccount4.BoundText = "", Null, DcAccount4.BoundText)

    
    rs("CurrencyID").value = IIf(val(Me.DcCurrency.BoundText) = 0, 1, (Me.DcCurrency.BoundText))
    
   rs("sub_contractor_Account").value = IIf(DcAccount3.BoundText = "", Null, Trim(DcAccount3.BoundText))
    rs("sub_contractor_name").value = IIf(DcAccount4.text = "", Null, DcAccount4.text)
    
    rs("prifix").value = IIf(Trim$(Me.DCPreFix.text) = "", Null, Trim$(Me.DCPreFix.text))
    rs("code").value = IIf(Trim$(Me.txtid.text) = "", Null, Trim$(Me.txtid.text))
    
    rs("Fullcode").value = IIf(Me.DCPreFix.text & Me.txtid.text = "", Null, Me.DCPreFix.text & Me.txtid.text)
    
    rs("Project_name").value = IIf(Trim$(Me.TXTprojectname.text) = "", Null, Trim$(Me.TXTprojectname.text))
    rs("Project_namee").value = IIf(Trim$(Me.TXTprojectnamee.text) = "", Null, Trim$(Me.TXTprojectnamee.text))
    
    rs("project_cost").value = IIf(val(Me.TxtProjectCosts.text) = 0, 0, val(Me.TxtProjectCosts.text))
    rs("general_discount").value = IIf(val(Me.txt_total_discount.text) = 0, 0, val(Me.txt_total_discount.text))
    
     rs("DiscountPercentage").value = IIf(val(Me.TxtDiscountPercentage.text) = 0, 0, val(Me.TxtDiscountPercentage.text))
     
    rs("cost_after_discount").value = rs("project_cost").value - rs("general_discount").value
    rs("net").value = rs("project_cost").value - rs("general_discount").value
     rs("Dept_ID").value = IIf(Trim$(Me.DcbDept.BoundText) = "", Null, Trim$(Me.DcbDept.BoundText))
    rs("branch_no").value = IIf(Trim$(Me.Dcbranch.BoundText) = "", Null, Trim$(Me.Dcbranch.BoundText))

    rs("Contract_type").value = IIf(Me.DataCombo5.BoundText = "", Null, Me.DataCombo5.BoundText)
    rs("Contract_type_name").value = IIf(Trim$(Me.DataCombo5.text) = "", Null, Trim$(Me.DataCombo5.text))
    rs("Project_status").value = IIf(Trim$(Me.DataCombo1.text) = "", Null, Trim$(Me.DataCombo1.text))
  
    '   Rs("departement").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    '   Rs("project_code").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
 
    rs("branch_no").value = IIf(Not IsNumeric(Dcbranch.BoundText), 0, Dcbranch.BoundText)
    
   ' rs("End_user_id").value = IIf(Trim$(Me.DCAccount1.text) = "", Null, Trim$(Me.DCAccount1.text))
   ' rs("sub_contractor_id").value = IIf(Trim$(Me.DCAccount3.text) = "", Null, Trim$(Me.DCAccount3.text))
    '   Rs("total").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    '  Rs("sub_discount_total").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    '  Rs("net").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    rs("items_total").value = IIf(Me.XPTxtSum.text = "", 0, Me.XPTxtSum.text)
  
  If Ptype(0).value = True Then
  
    rs("Pstate").value = 0
  Else
     rs("Pstate").value = 1
  End If
  
  
    If Me.OptType(2).value = True Then
        rs("OpenBalance").value = 0
        rs("OpenBalanceType").value = Null
    ElseIf Me.OptType(0).value = True Then
        rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
        rs("OpenBalanceType").value = 0
    ElseIf Me.OptType(1).value = True Then
        rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
        rs("OpenBalanceType").value = 1
    End If
    
    If Me.OptType1(2).value = True Then
        rs("OpenBalance1").value = 0
        rs("OpenBalanceType1").value = Null
    ElseIf Me.OptType1(0).value = True Then
        rs("OpenBalance1").value = val(Me.TxtOpenBalance1.text)
        rs("OpenBalanceType1").value = 0
    ElseIf Me.OptType1(1).value = True Then
        rs("OpenBalance1").value = val(Me.TxtOpenBalance1.text)
        rs("OpenBalanceType1").value = 1
    End If
    
    If Me.OptType2(2).value = True Then
        rs("OpenBalance2").value = 0
        rs("OpenBalanceType2").value = Null
    ElseIf Me.OptType2(0).value = True Then
        rs("OpenBalance2").value = val(Me.TxtOpenBalance2.text)
        rs("OpenBalanceType2").value = 0
    ElseIf Me.OptType2(1).value = True Then
        rs("OpenBalance2").value = val(Me.TxtOpenBalance2.text)
        rs("OpenBalanceType2").value = 1
    End If
    
    If Me.OptType3(2).value = True Then
        rs("OpenBalance3").value = 0
        rs("OpenBalanceType3").value = Null
    ElseIf Me.OptType3(0).value = True Then
        rs("OpenBalance3").value = val(Me.TxtOpenBalance3.text)
        rs("OpenBalanceType3").value = 0
    ElseIf Me.OptType3(1).value = True Then
        rs("OpenBalance3").value = val(Me.TxtOpenBalance3.text)
        rs("OpenBalanceType3").value = 1
    End If
 
    If Me.OptType4(2).value = True Then
        rs("OpenBalance4").value = 0
        rs("OpenBalanceType4").value = Null
    ElseIf Me.OptType4(0).value = True Then
        rs("OpenBalance4").value = val(Me.TxtOpenBalance4.text)
        rs("OpenBalanceType4").value = 0
    ElseIf Me.OptType4(1).value = True Then
        rs("OpenBalance4").value = val(Me.TxtOpenBalance4.text)
        rs("OpenBalanceType4").value = 1
    End If
    
    rs("OpenBalanceDate").value = Me.Dtp.value
    Dim Account_Code_dynamic1 As String

    If val(TxtOpenBalance.text) <> 0 Or val(TxtOpenBalance1.text) <> 0 Or val(TxtOpenBalance2.text) <> 0 Or val(TxtOpenBalance3.text) <> 0 Or val(TxtOpenBalance4.text) <> 0 Then
        
        Account_Code_dynamic1 = get_account_code_branch(73, my_branch)
                    
        If Account_Code_dynamic1 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
        Else

            If Account_Code_dynamic1 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                GoTo ErrTrap
                                 
            End If
        End If

        txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
        rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
    Else
        rs("opening_balance_voucher_id").value = Null
    End If
  
    'OPENING Balance Voucher
    Dim StrDes As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrDes = "«·—’Ìœ «·≈›  «ÕÏ ·‹ " & Trim(Me.TXTprojectname.text) & " "
    Else
        StrDes = " Opening Balance For: " & Trim(Me.TXTprojectnamee.text) & " "
    End If
        
    If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            Dim LngDevID As Long
            Dim LngOpenID As Long

            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType(0).value = True Then
        
                If ModAccounts.AddNewDev(LngDevID, 1, rs("expanses_account").value, val(Me.TxtOpenBalance.text), 0, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 1, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
            ElseIf Me.OptType(1).value = True Then
                 
                If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance.text), 0, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, rs("expanses_account").value, val(Me.TxtOpenBalance.text), 1, StrDes & " - ··„’—Ê›«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
        End If
    End If

    If Me.OptType1(0).value = True Or Me.OptType1(1).value = True Then
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType1(0).value = True Then
        
                If ModAccounts.AddNewDev(LngDevID, 1, rs("REVENUE_account").value, val(Me.TxtOpenBalance1.text), 0, StrDes & " - ··«Ì—«œ«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance1.text), 1, StrDes & " - ··«Ì—«œ«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
            ElseIf Me.OptType1(1).value = True Then
                 
                If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance1.text), 0, StrDes & " - ··«Ì—«œ«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, rs("REVENUE_account").value, val(Me.TxtOpenBalance1.text), 1, StrDes & " - ··«Ì—«œ«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
        End If
    End If
 
    If Me.OptType2(0).value = True Or Me.OptType2(1).value = True Then
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType2(0).value = True Then
        
                If ModAccounts.AddNewDev(LngDevID, 1, rs("Material_account").value, val(Me.TxtOpenBalance2.text), 0, StrDes & " - ··„Ê«œ ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance2.text), 1, StrDes & " - ··„Ê«œ ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
            ElseIf Me.OptType2(1).value = True Then
                 
                If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance2.text), 0, StrDes & " - ··„Ê«œ ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, rs("Material_account").value, val(Me.TxtOpenBalance2.text), 1, StrDes & " - ··„Ê«œ ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
        End If
    End If
  
    If Me.OptType3(0).value = True Or Me.OptType3(1).value = True Then
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType3(0).value = True Then
        
                If ModAccounts.AddNewDev(LngDevID, 1, rs("Salary_account").value, val(Me.TxtOpenBalance3.text), 0, StrDes & " - ··«ÃÊ— ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance3.text), 1, StrDes & " - ··«ÃÊ— ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
            ElseIf Me.OptType3(1).value = True Then
                 
                If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance3.text), 0, StrDes & " - ··«ÃÊ— ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, rs("Salary_account").value, val(Me.TxtOpenBalance3.text), 1, StrDes & " - ··«ÃÊ— ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
        End If
    End If
   
    If Me.OptType4(0).value = True Or Me.OptType4(1).value = True Then
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

            LngOpenID = 1
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
            If Me.OptType4(0).value = True Then
        
                If ModAccounts.AddNewDev(LngDevID, 1, rs("legal").value, val(Me.TxtOpenBalance4.text), 0, StrDes & " - ··„” Œ·’«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance4.text), 1, StrDes & " - ··„” Œ·’«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
            ElseIf Me.OptType4(1).value = True Then
                 
                If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance4.text), 0, StrDes & " - ··„” Œ·’«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
                
                If ModAccounts.AddNewDev(LngDevID, 2, rs("legal").value, val(Me.TxtOpenBalance4.text), 1, StrDes & " - ··„” Œ·’«  ", LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.text)) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
        End If
    End If
  
    'OPENING Balance Voucher
  
    rs.update
    ' If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
    ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
       
    ' »‰Êœ «·„‘—Ê⁄
    Set RsDev = New ADODB.Recordset
    RsDev.Open "projects_des", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    Dim i As Integer
    'Dim ExpensesID As Double
Dim Pand As Integer
    With Fg_Journal

        For i = .FixedRows To .Rows - 2

            '        Dim IntDEV_Type As Integer
            '        Dim SngDEV_Value As Single
            If .TextMatrix(i, .ColIndex("des")) <> "" Then

                RsDev.AddNew

                If .TextMatrix(i, .ColIndex("fullcode")) = "" Then
                    RsDev("fullcode").value = Me.txt_project_id & "-" & .TextMatrix(i, .ColIndex("LineNo"))
                Else
                    RsDev("fullcode").value = .TextMatrix(i, .ColIndex("fullcode"))
                End If
        If Me.TxtModFlg.text = "E" Then
        Pand = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("oprid"))), 0, .TextMatrix(i, .ColIndex("oprid")))
        
        If Me.Checked(Pand, 0) = True Then
        Else
       Pand = 1
        maxx Pand, 0
        End If
        End If
        If Me.TxtModFlg.text = "N" Then
         Pand = 1
    maxx Pand, 0
        End If
        RsDev("oprid").value = Pand
                RsDev("project_id").value = Me.txt_project_id.text
                RsDev("des").value = IIf(.TextMatrix(i, .ColIndex("des")) = "", Null, .TextMatrix(i, .ColIndex("des")))
                RsDev("qty").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("qty"))), 0, .TextMatrix(i, .ColIndex("qty")))
                RsDev("cost").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("cost"))), 0, .TextMatrix(i, .ColIndex("cost")))
                RsDev("total").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("total"))), 0, .TextMatrix(i, .ColIndex("total")))
                RsDev("discount").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("discount"))), 0, .TextMatrix(i, .ColIndex("discount")))
                RsDev("net").value = RsDev("total").value - RsDev("discount").value
                RsDev("line_no").value = .TextMatrix(i, .ColIndex("LineNo"))
                RsDev("sub_contractor_id").value = val(.TextMatrix(i, .ColIndex("sub_contractor_id")))
        
                RsDev.update
            End If

        Next i
    
    End With

    Set RsDevsub = New ADODB.Recordset
    RsDevsub.Open "Projectssub", Cn, adOpenStatic, adLockOptimistic, adCmdTable

'œ›⁄«  «·„‘—Ê⁄
 With GridSub

        For i = .FixedRows To .Rows - 2

            '        Dim IntDEV_Type As Integer
            '        Dim SngDEV_Value As Single
            If .TextMatrix(i, .ColIndex("id")) <> "" Then

                RsDevsub.AddNew

        
        
                RsDevsub("projectid").value = Me.txt_project_id
                RsDevsub("subdate").value = IIf(Not IsDate(.TextMatrix(i, .ColIndex("subdate"))), Null, .TextMatrix(i, .ColIndex("subdate")))
              '  RsDevsub("DesTerm").value = IIf(IsNull(.TextMatrix(i, .ColIndex("DesTerm"))), "", .TextMatrix(i, .ColIndex("DesTerm")))
                RsDevsub("ValueTerm").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("ValueTerm"))), 0, .TextMatrix(i, .ColIndex("ValueTerm")))
                RsDevsub("SubValue").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("SubValue"))), 0, .TextMatrix(i, .ColIndex("SubValue")))
                 RsDevsub("rate").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("rate"))), 0, .TextMatrix(i, .ColIndex("rate")))
                RsDevsub("REmarks").value = IIf(IsNull(.TextMatrix(i, .ColIndex("REmarks"))), "", .TextMatrix(i, .ColIndex("REmarks")))
      RsDevsub.update
            End If

        Next i
    
    End With

    'Retrive
    Command1(1).Enabled = False

    '„Ê«œ «·„‘—Ê⁄
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '       StrSQL = "Delete From Transaction_Details Where project_id =" & Val(Me.txt_project_id.text)
    '       Cn.Execute StrSQL, , adExecuteNoRecords
        
    '   Set RSTransDetails = New ADODB.Recordset
    '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '       For RowNum = 1 To FG.Rows - 1
    '       If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
  
    '   RSTransDetails.AddNew
    '   RSTransDetails("Transaction_Details").value = Val(txt_project_id.text)
    '   RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
    '   RSTransDetails("Price").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
    '   RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
    '   RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
    '
    '   RSTransDetails("UnitID").value = _
    '        IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
    '
    '   RSTransDetails.update
    '   End If
    '   Next
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '«·„ÊŸ›Ì‰ «·„”Ã·Ì‰ ›Ì «·„‘—Ê⁄
    'EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
 
    ' Set RsDev = New ADODB.Recordset
    ' Dim Sql As String
    '
    '  With VSFlexGrid1
    '    For I = .FixedRows To .Rows - 2
    '
    '        If .TextMatrix(I, .ColIndex("id")) <> "" Then
    '        Sql = "Select * from TblEmployee where Emp_ID=" & .TextMatrix(I, .ColIndex("id"))
    '        RsDev.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '        RsDev("project_id").value = Me.txt_project_id
    '        RsDev.update
    '        RsDev.Close
    '        End If
    '    Next I
    '
    'End With
    saveOperationDates val(txt_project_id.text), DTStartDate.value

    'EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then

                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„‘—Ê⁄" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " This Project Data Was Saved" & Chr(13)
                Msg = Msg + "Do you want To enter Another Project"
            End If

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Command1_Click (0)
                Exit Sub
            End If
            
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Amendments have been saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If

    End Select
    
    'If SystemOptions.UserInterface = EnglishInterface Then
    'MsgBox "Saved", vbInformation, ""
    'Else
    'MsgBox " „ Õ›Ÿ «·»Ì«‰« ", vbInformation, ""
    'End If
    TxtModFlg.text = "R"
    Exit Sub
ErrTrap:

    If SystemOptions.UserInterface = EnglishInterface Then
        MsgBox "Error During saving", vbInformation, ""
    Else
        MsgBox "ÕœÀ Œÿ√ «À‰«¡  Õ›Ÿ «·»Ì«‰« ", vbInformation, ""
    End If

End Sub

Function calcnets()

    With Me.VSFlexGrid1
        txt_employee_count = .Rows - 2
        Me.txt_emp_salary.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
    End With
 
End Function

  Private Sub ReLineGrid2()
    IntCounter = 0
    Dim SUM As Double
    Dim i As Integer
  SUM = 0
    With GridSub

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("subdate")) <> "" Then
            SUM = SUM + val(.TextMatrix(i, .ColIndex("SubValue")))
            IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("id")) = IntCounter
            If SUM <= val(total_after_discount.text) Then
                
                Else
                MsgBox "⁄œœ «·œ›⁄«  «ﬂ»— „‰ ﬁÌ„… «·„‘—Ê⁄"
                .TextMatrix(i, .ColIndex("SubValue")) = 0
                Exit Sub
                End If
  
            End If

        Next i
   
    End With
End Sub
Private Sub ReLineGrid(Optional current_terms As String = "")
    Dim i As Integer
    Dim IntCounter As Integer
    Dim StartWeek As Double
    Dim EndWeek As Double
    Dim EarlyStartWeek As Double
    Dim EarlyEndWeek As Double
    Dim rs As ADODB.Recordset

    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("des")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                .TextMatrix(i, .ColIndex("discount")) = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("discount"))), 0, .TextMatrix(i, .ColIndex("discount")))
                .TextMatrix(i, .ColIndex("qty")) = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("qty"))), 0, .TextMatrix(i, .ColIndex("qty")))
                .TextMatrix(i, .ColIndex("cost")) = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("cost"))), 0, .TextMatrix(i, .ColIndex("cost")))
                sql = "select sum(total) as total  From terms_operations Where term_fullcode='" & .TextMatrix(i, .ColIndex("fullcode")) & "'"
                Set rs = New ADODB.Recordset
                rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If rs.RecordCount > 0 And Not IsNull(rs("total").value) Then
                    .TextMatrix(i, .ColIndex("qty")) = ""
                    .TextMatrix(i, .ColIndex("cost")) = ""
                    .TextMatrix(i, .ColIndex("total")) = rs("total").value
                    .TextMatrix(i, .ColIndex("net")) = rs("total").value
         
                Else
                    .TextMatrix(i, .ColIndex("total")) = .TextMatrix(i, .ColIndex("qty")) * .TextMatrix(i, .ColIndex("cost"))
                    .TextMatrix(i, .ColIndex("net")) = .TextMatrix(i, .ColIndex("total")) - .TextMatrix(i, .ColIndex("discount"))
         
                End If

                .TextMatrix(i, .ColIndex("fullcode")) = Me.txt_project_id.text & "-" & IntCounter
            End If

        Next i

        Me.txt_total_sum.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        Me.txt_sub_discount.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("discount"), .Rows - 1, .ColIndex("discount"))
        Me.txt_sub_net.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("net"), .Rows - 1, .ColIndex("net"))
    End With
       
    Label13.Caption = getoprTitle
    IntCounter = 0

    With VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
  
            End If

        Next i
   
        txt_employee_count = .Rows - 1
        Me.txt_emp_salary.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))

    End With
     
    IntCounter = 0

    With VSFlexGrid2

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                '.TextMatrix(i, .ColIndex("fullcode")) = current_terms & "-" & IntCounter
                .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("total_expenses"))) + val(.TextMatrix(i, .ColIndex("total_items"))) + val(.TextMatrix(i, .ColIndex("total_salary")))

            End If
        
            If val(.TextMatrix(i, .ColIndex("Period1"))) = 0 Then
                .TextMatrix(i, .ColIndex("Critical")) = 1
            Else
                .TextMatrix(i, .ColIndex("Critical")) = 0
            End If

            get_opr_details .TextMatrix(i, .ColIndex("Pre")), val(.TextMatrix(i, .ColIndex("period"))), val(.TextMatrix(i, .ColIndex("period1"))), StartWeek, EndWeek, EarlyStartWeek, EarlyEndWeek
            .TextMatrix(i, .ColIndex("startweek")) = StartWeek
            .TextMatrix(i, .ColIndex("EndWeek")) = EndWeek
            .TextMatrix(i, .ColIndex("Earlystartweek")) = EarlyStartWeek
            .TextMatrix(i, .ColIndex("EarlyEndWeek")) = EarlyEndWeek

            If val(.TextMatrix(i, .ColIndex("Critical"))) Then
                .Cell(flexcpBackColor, i, 14, i, 14) = vbRed
            Else
                .Cell(flexcpBackColor, i, 14, i, 14) = vbGreen
            End If

        Next i

        Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        Me.TXTNoOFWeek.text = .Aggregate(flexSTMax, .FixedRows, .ColIndex("EndWeek"), .Rows - 1, .ColIndex("EndWeek"))
        Dim X As Double
        X = getProjectDuration(val(Me.txt_project_id.text))
        Text4.text = X & " " & Me.Label13.Caption
 
        If SystemOptions.ProcessPeriodType = 0 Then
            DTEnddate.value = DateAdd("d", X, DTStartDate.value)      'day
        ElseIf SystemOptions.ProcessPeriodType = 1 Then
            DTEnddate.value = DateAdd("m", X, DTStartDate.value)   'Month
        ElseIf SystemOptions.ProcessPeriodType = 2 Then
            DTEnddate.value = DateAdd("yyyy", X, DTStartDate.value)   'Year
        ElseIf SystemOptions.ProcessPeriodType = 3 Then
            DTEnddate.value = DateAdd("ww", X, DTStartDate.value)   'week
        End If

    End With

    IntCounter = 0

    With VSFlexGrid3

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("ExpensesID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

Private Sub REFillOprData(TblProcessDEFID As Double)
    Dim i As Integer
    Dim IntCounter As Integer
    Dim StartWeek As Double
    Dim EndWeek As Double
    Dim EarlyStartWeek As Double
    Dim EarlyEndWeek As Double
    Dim rs As ADODB.Recordset

    With VSFlexGrid2

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("OPRIDD")) <> "" Then
                sql = "   SELECT      dbo.TblProcessDEF.*, dbo.TblProcessUnites.UnitName, dbo.TblProcessUnites.UnitNamee "
                sql = sql & "  from dbo.TblProcessDEF  "
                sql = sql & "   INNER JOIN"
                sql = sql & " dbo.TblProcessUnites ON dbo.TblProcessDEF.UnitID = dbo.TblProcessUnites.UnitID"
                sql = sql & " Where (TblProcessDEFID = " & TblProcessDEFID & ")"
                                  
                Set rs = New ADODB.Recordset
                rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If rs.RecordCount > 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("unitname")) = IIf(IsNull(rs("unitname").value), "", rs("unitname").value)
                Else
                .TextMatrix(i, .ColIndex("unitname")) = IIf(IsNull(rs("unitnamee").value), "", rs("unitnamee").value)
                End If
                
                    Dim IntervalID As Integer
                    Dim intervaltype As String
                    IntervalID = IIf(IsNull(rs("Intervalid").value), 0, rs("Intervalid").value)
If SystemOptions.UserInterface = ArabicInterface Then
                    If IntervalID = 0 Then
                        intervaltype = "œﬁÌﬁ…"
                    ElseIf IntervalID = 1 Then
                        intervaltype = "”«⁄Â"
                    ElseIf IntervalID = 2 Then
                        intervaltype = "ÌÊ„"
                    ElseIf IntervalID = 3 Then
                        intervaltype = "«”»Ê⁄"
                    ElseIf IntervalID = 4 Then
                        intervaltype = "‘Â—"
                    ElseIf IntervalID = 5 Then
                        intervaltype = "”‰Â"
                    End If
Else
               If IntervalID = 0 Then
                        intervaltype = "Minute"
                    ElseIf IntervalID = 1 Then
                        intervaltype = "Hour"
                    ElseIf IntervalID = 2 Then
                        intervaltype = "Day"
                    ElseIf IntervalID = 3 Then
                        intervaltype = "Week"
                    ElseIf IntervalID = 4 Then
                        intervaltype = "Month"
                    ElseIf IntervalID = 5 Then
                        intervaltype = "Year"
                    End If
End If

                    .TextMatrix(i, .ColIndex("period")) = IIf(IsNull(rs("interval").value), 0, rs("interval").value) * val(.TextMatrix(i, .ColIndex("qty")))
                    .TextMatrix(i, .ColIndex("periodView")) = .TextMatrix(i, .ColIndex("period")) & "  " & intervaltype
                                                         
                End If
                             
            End If

        Next i
  
    End With

End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "id='" & val(txt_project_id.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If txt_project_id.text <> "" Then
        StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where project_id=" & val(txt_project_id.text)
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

       If Not (RsTemp.EOF Or RsTemp.BOF) Then
             Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·„‘—Ê⁄" & Chr(13)
              Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «·„‘—Ê⁄"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
             Exit Sub
         End If

        Msg = "”Ì „ Õ–› »Ì«‰«  «·„‘—Ê⁄ —ﬁ„ " & Chr(13)
        Msg = Msg + (txt_project_id.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                Dim StrAccountCode As String
                StrAccountCode = rs("expanses_account").value

                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
               
                Else
                    Exit Sub
                End If
            
                StrAccountCode = rs("REVENUE_account").value

                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
               
                Else
                    Exit Sub
                End If
            
                StrAccountCode = rs("Material_account").value

                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
               
                Else
                    Exit Sub
                End If
            
                StrAccountCode = rs("Salary_account").value

                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
               
                Else
                    Exit Sub
                End If
            
                StrAccountCode = rs("legal").value

                If ModAccounts.DeleteAccount(StrAccountCode) = True Then
                
                Else
                    Exit Sub
                End If
                                
                StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblExpensiveOper Where ProjectID=" & val(Me.txt_project_id.text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblMatrials Where ProjectID=" & val(Me.txt_project_id.text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblEmpOper Where ProjectID=" & val(Me.txt_project_id.text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblEquepment Where ProjectID=" & val(Me.txt_project_id.text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
             
                rs.delete
            
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                
                    Fg_Journal.Clear flexClearScrollable, flexClearEverything
                    Fg_Journal.Rows = 2
                    Fg_Journal.Enabled = True
                    txt_total_discount = 0
          
                    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                    VSFlexGrid1.Rows = 2
                    VSFlexGrid1.Enabled = True
          
                    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
                    VSFlexGrid2.Rows = 2
                    VSFlexGrid2.Enabled = True
          
                    '    XPTxtCurrent.Caption = 0
                    '    XPTxtCount.Caption = 0
                Else
                    XPBtnMove_Click (0)
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·Œ“‰… "
        Msg = Msg & Chr(13) & Err.description
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If
 
End Sub
 


Private Sub Cmd_Click()
RemoveGridRow
End Sub

Private Sub CMDViewGantt_Click(Index As Integer)
    terms_operations_Click (1)
    Gantt.show
  '  Gantt.Init_Chart val(TXTNoOFWeek.text)
   Gantt.Draw_Data current_terms
End Sub

Private Sub Command1_Click(Index As Integer)
    'On Error Resume Next
    Dim FirstPeriodDateInthisYear As Date

    Select Case Index

        Case 4
            Fra(2).Visible = False
            Fra(3).Visible = True

        Case 0
            'mod_flad = "N"

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If
        
            TxtModFlg.text = "N"
            clear_all Me
            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
            Me.Dtp = FirstPeriodDateInthisYear
            Me.Dtp1 = FirstPeriodDateInthisYear
            Me.Dtp2 = FirstPeriodDateInthisYear
            Me.Dtp3 = FirstPeriodDateInthisYear
            Me.Dtp4 = FirstPeriodDateInthisYear

            OptType(2).value = True
            OptType1(2).value = True
            OptType2(2).value = True
            OptType3(2).value = True
            OptType4(2).value = True

            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 2
            Fg_Journal.Enabled = True
            txt_total_discount = 0 '
          
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 2
            VSFlexGrid1.Enabled = True
          
            VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.Rows = 2
            VSFlexGrid2.Enabled = True
          
                  GridSub.Clear flexClearScrollable, flexClearEverything
            GridSub.Rows = 2
            GridSub.Enabled = True
            
            Option4.value = True

            Me.txt_project_id.text = CStr(new_id("projects", "id", "", True))
            Command1(1).Enabled = True
            Me.Dcbranch.BoundText = Current_branch
            XPDtbBill.value = Date
Ptype(0).value = True

        Case 1
            SaveData
            my_branch = Me.Dcbranch.BoundText

            'Fg_Journal.Enabled = False
        Case 2 '

            On Error Resume Next

            If SystemOptions.UserInterface = EnglishInterface Then
                If DCPreFix.text & txtid.text = "" Then MsgBox "Select Project firstly": Exit Sub

            Else

                If DCPreFix.text & txtid.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „‘—Ê⁄ «Ê·«": Exit Sub

            End If

            imaged.show

            If SystemOptions.UserInterface = EnglishInterface Then

                imaged.Label9.Caption = "Attachment For Project "
                imaged.Caption = "Project Attachment  "
                imaged.Label6.Caption = "   Project NO"
                Label5.Caption = "Documents"

            Else

                imaged.Label9.Caption = "„—›ﬁ«    „‘—Ê⁄  —ﬁ„"
                imaged.Caption = "„—›ﬁ«  „‘—Ê⁄  "
                imaged.Label6.Caption = "—ﬁ„  «·„‘—Ê⁄"

            End If

            imaged.SUBJECT_NO = DCPreFix.text & txtid.text
            imaged.txtopeation_type = "„—›ﬁ«  „‘—Ê⁄"
 
            imaged.Adodc1.ConnectionString = Cn.ConnectionString
            imaged.Adodc1.CommandType = adCmdText
            imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = '„—›ﬁ«  „‘—Ê⁄' and subject_no='" & DCPreFix.text & txtid.text & "'"
            imaged.Adodc1.Refresh

            If imaged.Adodc1.Recordset.RecordCount > 0 Then

                imaged.DBPix201.Visible = True
            Else
                imaged.DBPix201.Visible = False
            End If

        Case 3
 
            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If
        
            TxtModFlg.text = "E"
            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
            Me.Dtp = FirstPeriodDateInthisYear
            Me.Dtp1 = FirstPeriodDateInthisYear
            Me.Dtp2 = FirstPeriodDateInthisYear

            Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
           
            VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
            VSFlexGrid1.Enabled = True
            VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
            VSFlexGrid2.Enabled = True
            VSFlexGrid3.Rows = VSFlexGrid3.Rows + 1
            VSFlexGrid3.Enabled = True
                     GridSub.Rows = GridSub.Rows + 1
            GridSub.Enabled = True
            Command1(1).Enabled = True

            'SaveData
        Case 4
 
        Case 5

   
            FrmProjectSearch.show vbModal


        Case 6
            Undo

        Case 7

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If
        
            Del_Trans
            
    Case 8
    Fra(2).Visible = True
    Fra(3).Visible = False
            
    End Select

End Sub

Function create_accounts() As Boolean
    Dim RsSavRec As ADODB.Recordset
    Dim RsSavRec1 As ADODB.Recordset
    Dim RsSavRec2 As ADODB.Recordset
    Dim RsSavRec3 As ADODB.Recordset
    Dim RsSavRec4 As ADODB.Recordset
    Dim RsSavRec5 As ADODB.Recordset

    Dim My_SQL As String

    If 1 = 1 Then
         
        Account_Code_dynamic1 = get_account_code_branch(14, my_branch)
        
        If Account_Code_dynamic1 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic1 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«» „’—Ê›«   ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        
                create_accounts = False
                Exit Function
            End If
        End If
        
        Account_Code_dynamic2 = get_account_code_branch(15, my_branch)
        
        If Account_Code_dynamic2 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic2 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«» «Ì—«œ«  ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        
                create_accounts = False
                Exit Function
            End If
        End If
        
        Account_Code_dynamic3 = get_account_code_branch(27, my_branch)
        
        If Account_Code_dynamic3 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic3 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«» „Ê«œ  ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        
                create_accounts = False
                Exit Function
            End If
        End If
        
        Account_Code_dynamic4 = get_account_code_branch(28, my_branch)
        
        If Account_Code_dynamic4 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic4 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«» «ÃÊ— ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        
                create_accounts = False
                Exit Function
            End If
        End If
        
        Account_Code_dynamic5 = get_account_code_branch(32, my_branch)
        
        If Account_Code_dynamic5 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        
            create_accounts = False
            Exit Function
        Else

            If Account_Code_dynamic5 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«» ‰Ÿ«„Ì ··„‘«—Ì⁄ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        
                create_accounts = False
                Exit Function
            End If
        End If
 
        Account_Code_dynamic1C = ModAccounts.AddNewAccount(Account_Code_dynamic1, TXTprojectname & " -„’—Ê›«  ", True, False, TXTprojectnamee & " -EXPANSES")
        Account_Code_dynamic2C = ModAccounts.AddNewAccount(Account_Code_dynamic2, TXTprojectname & "-«Ì—«œ«  ", True, False, TXTprojectnamee & " -REVENUE")
        Account_Code_dynamic3C = ModAccounts.AddNewAccount(Account_Code_dynamic3, TXTprojectname & " -„Ê«œ  ", True, False, TXTprojectnamee & " -Material ")
        Account_Code_dynamic4C = ModAccounts.AddNewAccount(Account_Code_dynamic4, TXTprojectname & " -⁄„«·Â ", True, False, TXTprojectnamee & " -salary")
        Account_Code_dynamic5C = ModAccounts.AddNewAccount(Account_Code_dynamic5, TXTprojectname & " -„” Œ·’«  ", True, False, TXTprojectnamee & " -legal")
        create_accounts = True

        Exit Function

    End If
       create_accounts = True

        Exit Function
        
    My_SQL = "  select * from project_category WHERE branch_id =" & Dcbranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"

    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec.RecordCount = 0 Then MsgBox "this project type not found in this branch", vbCritical: create_accounts = False: Exit Function

    My_SQL = "  select * from project_category WHERE type='a14' and branch_id =" & Dcbranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"
    Set RsSavRec1 = New ADODB.Recordset
    RsSavRec1.CursorLocation = adUseClient
    RsSavRec1.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec1.RecordCount = 0 Then MsgBox "·„ Ì „  ÕœÌœ Õ”«» „’—Ê›«  ·Â–« «·„‘—Ê⁄", vbCritical: create_accounts = False: Exit Function
 
    My_SQL = "  select * from project_category WHERE type='a15' and branch_id =" & Dcbranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"

    Set RsSavRec2 = New ADODB.Recordset
    RsSavRec2.CursorLocation = adUseClient
    RsSavRec2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec2.RecordCount = 0 Then MsgBox "·„ Ì „  ÕœÌœ Õ”«» «Ì—«œ«  ·Â–« «·„‘—Ê⁄", vbCritical: create_accounts = False:  Exit Function

    My_SQL = "  select * from project_category WHERE type='a27' and branch_id =" & Dcbranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"
    Set RsSavRec3 = New ADODB.Recordset
    RsSavRec3.CursorLocation = adUseClient
    RsSavRec3.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec3.RecordCount = 0 Then MsgBox "·„ Ì „  ÕœÌœ Õ”«» „Ê«œ ·Â–« «·„‘—Ê⁄", vbCritical:  create_accounts = False: Exit Function
 
    My_SQL = "  select * from project_category WHERE type='a28' and branch_id =" & Dcbranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"

    Set RsSavRec4 = New ADODB.Recordset
    RsSavRec4.CursorLocation = adUseClient
    RsSavRec4.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec4.RecordCount = 0 Then MsgBox "·„ Ì „  ÕœÌœ Õ”«» „Ê«œ ·Â–« «·„‘—Ê⁄", vbCritical:  create_accounts = False: Exit Function
 
    My_SQL = "  select * from project_category WHERE type='a32' and branch_id =" & Dcbranch.BoundText & "and category_id=" & DataCombo5.BoundText & " order by id"

    Set RsSavRec5 = New ADODB.Recordset
    RsSavRec5.CursorLocation = adUseClient
    RsSavRec5.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsSavRec5.RecordCount = 0 Then MsgBox "·„ Ì „  ÕœÌœ Õ”«» ‰Ÿ«„Ì ·Â–« «·„‘—Ê⁄", vbCritical:  create_accounts = False: Exit Function
 
    Account_Code_dynamic1C = ModAccounts.AddNewAccount(RsSavRec1("account_code").value, DCPreFix.text & Trim$(Me.txtid.text) & " -„’—Ê›«  ", True, False, DCPreFix.text & Trim$(Me.txtid.text) & " -Expenses")
    Account_Code_dynamic2C = ModAccounts.AddNewAccount(RsSavRec2("account_code").value, DCPreFix.text & Trim$(Me.txtid.text) & "  -«Ì—«œ«  ", True, False, DCPreFix.text & Trim$(Me.txtid.text) & "  -Revenue")
    Account_Code_dynamic3C = ModAccounts.AddNewAccount(RsSavRec3("account_code").value, DCPreFix.text & Trim$(Me.txtid.text) & "  -„Ê«œ ", True, False, DCPreFix.text & Trim$(Me.txtid.text) & "  -Material")
    Account_Code_dynamic4C = ModAccounts.AddNewAccount(RsSavRec4("account_code").value, DCPreFix.text & Trim$(Me.txtid.text) & "  -⁄„«·Â ", True, False, DCPreFix.text & Trim$(Me.txtid.text) & "  -Salary")
    Account_Code_dynamic5C = ModAccounts.AddNewAccount(RsSavRec5("account_code").value, DCPreFix.text & Trim$(Me.txtid.text) & "  -„” Œ·’«  ", True, False, DCPreFix.text & Trim$(Me.txtid.text) & "  -Legal")
  
    ' Me.legal.text = ModAccounts.AddNewAccount(RsSavRec4("account_code").value, DCPreFix.text & Trim$(Me.txtid.text) & "  Legal ", True, False)

    create_accounts = True

End Function
 
Private Sub Command2_Click()
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String

    If Me.dcJobTypeName.BoundText <> "" Then
        If Not IsNumeric(TxtCount.text) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ  ⁄œœ «·«Ì«„   "
            Else
                Msg = " SPecify No of Days  "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtCount.SetFocus
            Exit Sub
        End If
        
        If Not IsNumeric(TxtEmpcount.text) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ  ⁄œœ «·„ÿ·Ê»Ì‰ „‰ Â–… «·„Â‰…  "
            Else
                Msg = "Specify No oF labors From this Job  "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtCount.SetFocus
            Exit Sub
        End If

        If Option4.value = True Then ' ﬁœÌ— ›ﬁÿ
            StrSQL = "SELECT     ROUND((ISNULL(dbo.TblEmployee.Emp_Salary, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_sakn, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_bus, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_food, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_others, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_mob, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_mang, 0)) / 30, 2) AS daysalary, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, " & "       dbo.TblEmployee.JobTypeID , dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName " & "       FROM         dbo.TblEmployee INNER JOIN" & "       dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID " & "  WHERE      dbo.TblEmployee.JobTypeID =" & val(Me.dcJobTypeName.BoundText)
            ' StrSQL = "SELECT  round((isnull(Emp_Salary,0)+isnull(Emp_Salary_sakn,0) +isnull(Emp_Salary_bus,0)   +isnull(Emp_Salary_food,0)  +isnull(Emp_Salary_others,0)  +isnull(Emp_Salary_mob,0)  +isnull(Emp_Salary_mang,0))/30,2) as daysalary,* from TblEmployee Where  JobTypeID= " & Val(Me.dcJobTypeName.BoundText)
        ElseIf Option5.value = True Then
            ' StrSQL = "SELECT  round((isnull(Emp_Salary,0)+isnull(Emp_Salary_sakn,0) +isnull(Emp_Salary_bus,0)   +isnull(Emp_Salary_food,0)  +isnull(Emp_Salary_others,0)  +isnull(Emp_Salary_mob,0)  +isnull(Emp_Salary_mang,0))/30,2) as daysalary,* from TblEmployee Where  project_id=0 and JobTypeID= " & Val(Me.dcJobTypeName.BoundText) '   Œ’Ì’ ›⁄·Ï
            StrSQL = "SELECT     ROUND((ISNULL(dbo.TblEmployee.Emp_Salary, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_sakn, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_bus, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_food, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_others, 0) + ISNULL(dbo.TblEmployee.Emp_Salary_mob, 0) " & "       + ISNULL(dbo.TblEmployee.Emp_Salary_mang, 0)) / 30, 2) AS daysalary, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, " & "       dbo.TblEmployee.JobTypeID , dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName " & "       FROM         dbo.TblEmployee INNER JOIN " & "       dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID " & " WHERE      dbo.TblEmployee.JobTypeID =" & val(Me.dcJobTypeName.BoundText) & " and ( dbo.TblEmployee.project_id =0 OR  dbo.TblEmployee.project_id IS NULL)"
                
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Dim lastrow As Integer
        Dim X As Integer

        If rs.RecordCount > 0 Then
            If Option5.value = True Then
                If rs.RecordCount < val(TxtEmpcount.text) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "«·⁄œœ «·„ÿ·Ê» „‰ «·⁄„«· €Ì— „ Ê›— Â·  —Ìœ  «· ﬂ„·… »«·⁄œœ «·„ÊÃÊœ" & Chr(13)
                        Msg = Msg & "  ‰⁄„  ﬂ„·…"
                        Msg = Msg & "  ·«  «·€«¡" & Chr(13)
                    Else
                        Msg = "No Of Labors not exist Now,continue with avilable " & Chr(13)
                        Msg = Msg & "  Yes -continue  "
                        Msg = Msg & " No - cancel" & Chr(13)
                    End If

                    X = MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

                    If X = vbNo Then
                        Exit Sub
                    End If
                
                End If
            End If

            rs.MoveFirst
    
            With Me.VSFlexGrid1
                lastrow = .Rows - 1
                .Rows = .Rows + rs.RecordCount

                For i = lastrow To .Rows - 2
                
                    .TextMatrix(i, .ColIndex("LineNo")) = i
                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                    .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
                        
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                        
                    .TextMatrix(i, .ColIndex("jobid")) = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)
                    .TextMatrix(i, .ColIndex("jobname")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                    .TextMatrix(i, .ColIndex("daysalary")) = Round(IIf(IsNull(rs("daysalary").value), 0, rs("daysalary").value), 2)
                    .TextMatrix(i, .ColIndex("Count")) = val(Me.TxtCount.text)
                    .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("daysalary"))) * val(.TextMatrix(i, .ColIndex("Count")))
                    rs.MoveNext
                Next

            End With

            calcnets
        Else

            If Option4.value = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "€Ì— „ Ê›— ⁄„«· »Â–… «·„Â‰…  "
                Else
                    Msg = "No Labors assigned to this job  "
                End If

            ElseIf Option5.value = True Then

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "€Ì— „ Ê›— ⁄„«· »Â–… «·„Â‰… «Ê «‰ ﬂ· «·⁄„«· „Œ’’Ì‰ ·„‘«—Ì⁄ «Ê ⁄„·Ì«  «Œ—Ï  "
                Else
                    Msg = "No Labors assigned to this job Or all Labors Allocated to another Project Process  "
                End If
            End If
                     
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Õœœ «·„Â‰… «Ê·« «·„ÿ·Ê»… «Ê·« "
        Else
            Msg = "Specify Job Firstly "
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        dcJobTypeName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

End Sub

Public Function ShowReports()
    On Error Resume Next

    Dim My_SQL2 As String
    Dim NetExpensen As Double

    Dim Balance As String
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    
    rs2.Open "projects", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
 Dim last_account As Integer
 last_account = 1
 Dim StrSQL As String
  Dim Fromdate As Date
  Dim toDate As Date
  Fromdate = "01/01/2014"
 toDate = Date
 
     Dim openingbalacedate As Date
    getOpeningBalancedate , , , , year(toDate), openingbalacedate, True

  'StrSQL = StrSQL & " SET opening_balance= dbo.GetOpeningBalance('" & SQLDate(openingbalacedate) , Account_code,1)"



        StrSQL = " update projects"
     
        StrSQL = StrSQL & " set  expansesE_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(toDate) & "', expanses_account, last_account),"
    StrSQL = StrSQL & " set  expansesM_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(toDate) & "', Material_account, last_account),"
       StrSQL = StrSQL & " set  expansesS_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(toDate) & "', Salary_account, last_account),"
          StrSQL = StrSQL & " set  REVENUE_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(toDate) & "', REVENUE_account, last_account),"
          StrSQL = StrSQL & " set  Legal_account_balance= dbo.GetBalance('" & SQLDate(Fromdate) & "','" & SQLDate(toDate) & "', legal, last_account)+ dbo.GetOpeningBalance(  SQLDate(" & openingbalacedate & ") , legal,1) "
          
       
          
   Cn.Execute StrSQL
        
    Dim i As Integer

  '  For i = 1 To rs2.RecordCount
  '      NetExpensen = 0
  '      WriteCustomerBalPublic rs2("expanses_account"), Balance
  '      rs2("expansesE_account_balance").value = val(Balance)
  '      NetExpensen = NetExpensen + val(Balance)
'
'        WriteCustomerBalPublic rs2("Material_account"), Balance
'        rs2("expansesM_account_balance").value = val(Balance)
'        NetExpensen = NetExpensen + val(Balance)
'
'        WriteCustomerBalPublic rs2("Salary_account"), Balance
'        rs2("expansesS_account_balance").value = val(Balance)
'        NetExpensen = NetExpensen + val(Balance)
'
'        WriteCustomerBalPublic rs2("REVENUE_account"), Balance
'
'        rs2("REVENUE_account_balance").value = val(Balance)
'
'        WriteCustomerBalPublic rs2("legal"), Balance
'
     '   rs2("Legal_account_balance").value = val(Balance) '
''
 '       rs2("expanses_account_balance").value = NetExpensen
 
 '       rs2.update
 '       rs2.MoveNext
 '   Next i

    Dim xApp As New CRAXDRT.Application

    Dim EmpReport As ClsEmployeeReport
    Dim xReport As New CRAXDRT.Report

    Dim rs As ADODB.Recordset
    Dim cCompanyInfo As ClsCompanyInfo
    Set cCompanyInfo = New ClsCompanyInfo
    sql = "SELECT * from projects"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
       
    If SystemOptions.UserInterface = ArabicInterface Then
        Set xReport = xApp.OpenReport(App.path & "\reports\construction\REPORT1A.rpt")
    Else
    
        Set xReport = xApp.OpenReport(App.path & "\reports\construction\REPORT1.rpt")
    End If

    xReport.Database.SetDataSource rs
 
    Set FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
    FrmReport.txtpath = (App.path & "\reports\construction\REPORT1A.rpt")
    FrmReport.CRViewer.viewReport
 
    xReport.reporttitle = cCompanyInfo.ArabCompanyName
    FrmReport.show
    Screen.MousePointer = vbDefault
    '   xReport.ReportTitle = X
    SendKeys "{RIGHT}"

End Function

Private Sub Command3_Click()

    If DoPremis(Do_Print, Me.name, True) = False Then
        Exit Sub
    End If

    ShowReports
End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String

        My_SQL = "  select id,name from project_status  "
        fill_combo DataCombo1, My_SQL
    End If

End Sub

Private Sub DataCombo5_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
        My_SQL = "  select id,name from contract_type  "
        fill_combo DataCombo5, My_SQL
    End If

End Sub

Private Sub DcAccount1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = 13 Then

        On Error Resume Next

        If DcAccount1.text = "" Then DcAccount2.text = "": Exit Sub
        DcAccount2.text = ""
        Dim My_SQL As String

        My_SQL = "select CusName from TblCustemers where CusID='" & DcAccount1.text & "'"
 
        Set rec = New ADODB.Recordset
        rec.CursorLocation = adUseClient

        rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(rec.Fields("Account_Name").value) Then
            DcAccount2.text = rec.Fields("CusName").value
        Else
            DcAccount2.text = ""
 
        End If

    End If

 '   If KeyCode = vbKeyF5 Then
      
 '       My_SQL = "  select Account_Code,CusID from TblCustemers  where type=1 "
 '
 '       fill_combo DcAccount1, MySQL

    'End If
        
End Sub

Private Sub DcAccount2_Change()
Dim fullcode As String
Dim DefaultSalesPersonId As Integer
      fullcode = ""
        GetCustomersDetail val(DcAccount2.BoundText), DefaultSalesPersonId, fullcode
        TxtCustCode.text = fullcode
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

        
  

        DcEmp1.BoundText = DefaultSalesPersonId
    

    
    End If
End Sub

Private Sub DCAccount2_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 8
            FrmCustemerSearch.show vbModal
            
        End If
End Sub

Private Sub DcAccount3_Click(Area As Integer)
    On Error Resume Next

    If DcAccount3.text = "" Then Exit Sub
    DcAccount4.text = ""
    Dim My_SQL As String

    My_SQL = "select CusName from TblCustemers where CusID='" & DcAccount3.text & "'"
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not IsNull(rec.Fields("Account_Name").value) Then
        DcAccount4.text = rec.Fields("CusName").value
    Else
        DcAccount4.text = ""
    End If

End Sub

Private Sub DcAccount1_Click(Area As Integer)
    On Error Resume Next

    If DcAccount1.text = "" Then Exit Sub
    DcAccount2.text = ""
    Dim My_SQL As String

    My_SQL = "select CusName from TblCustemers where CusID='" & DcAccount1.text & "'"
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not IsNull(rec.Fields("Account_Name").value) Then
        DcAccount2.text = rec.Fields("CusName").value
    Else
        DcAccount2.text = ""
    End If

End Sub

Private Sub DcAccount2_Click(Area As Integer)
    On Error Resume Next

    If DcAccount2.text = "" Then Exit Sub
    DcAccount1.text = ""
    Dim My_SQL As String

    My_SQL = "select CusID from TblCustemers where CusID =" & DcAccount2.BoundText
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not IsNull(rec.Fields("CusID").value) Then
        DcAccount1.text = rec.Fields("CusID").value
    Else
        DcAccount1.text = ""
    End If



Dim fullcode As String
Dim DefaultSalesPersonId As Integer
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

        
        fullcode = ""
        GetCustomersDetail val(DcAccount2.BoundText), DefaultSalesPersonId, fullcode
        TxtCustCode.text = fullcode

        DcEmp1.BoundText = DefaultSalesPersonId
    

    
    End If




End Sub

Private Sub DcAccount3_KeyUp(KeyCode As Integer, _
                             Shift As Integer)
    On Error Resume Next

    If KeyCode = 13 Then

        If DcAccount3.text = "" Then DcAccount4.text = "": Exit Sub
        DcAccount4.text = ""
        Dim My_SQL As String

        My_SQL = "select CusName from TblCustemers where CusID='" & DcAccount3.text & "'"
        Dim rec As ADODB.Recordset
        Set rec = New ADODB.Recordset
        rec.CursorLocation = adUseClient

        rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not IsNull(rec.Fields("Account_Name").value) Then
            DcAccount4.text = rec.Fields("CusName").value
        Else
            DcAccount4.text = ""
             
        End If
 
    End If

    If KeyCode = vbKeyF5 Then
        My_SQL = "  select Account_Code,CusID from TblCustemers  where type=1 "
                
        fill_combo DcAccount3, My_SQL

    End If
        
End Sub

Private Sub DcAccount4_Change()
Dim fullcode As String
Dim DefaultSalesPersonId As Integer
 
        fullcode = ""
        GetCustomersDetail val(DcAccount4.BoundText), DefaultSalesPersonId, fullcode
        TxtCustCode1.text = fullcode

        
    

    
  
End Sub

Private Sub DcAccount4_Click(Area As Integer)
    On Error Resume Next

    If DcAccount4.text = "" Then Exit Sub
    DcAccount3.text = ""
    Dim My_SQL As String

    My_SQL = "select CusID from TblCustemers where CusID =" & DcAccount4.BoundText
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not IsNull(rec.Fields("CusID").value) Then
        DcAccount3.text = rec.Fields("CusID").value
    Else
        DcAccount3.text = ""
    End If

Dim fullcode As String
Dim DefaultSalesPersonId As Integer
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

        
        fullcode = ""
        GetCustomersDetail val(DcAccount4.BoundText), DefaultSalesPersonId, fullcode
        TxtCustCode1.text = fullcode

        
    

    
    End If




End Sub

Function gettotal(X As String, filed As String, table As String, filed_search As String) As Double
    Dim My_SQL As String

    My_SQL = "  select Sum(" & filed & ") as total  from " & table & " where " & filed_search & "='" & X & "'"
    Dim rec As ADODB.Recordset
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient

    rec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    gettotal = IIf(IsNull(rec.Fields("total").value), 0, rec.Fields("total").value)

End Function

Private Sub DcAccount4_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
          FrmCustemerSearch.SearchType = 9
             FrmCustemerSearch.show vbModal
          
        End If
End Sub

Private Sub DCboItemsCode_Click(Area As Integer)

    If val(DCboItemsCode.BoundText) = 0 Then Exit Sub
    Text6.text = get_item_Reserved_qty(val(DCboItemsCode.BoundText))
    Text3.text = get_item_qty(val(DCboItemsCode.BoundText))
    Text1.text = get_item_Order_qty(val(DCboItemsCode.BoundText))

End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches Dcbranch
    End If

End Sub

Private Sub DCCurrency_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String

        If SystemOptions.UserInterface = ArabicInterface Then
            My_SQL = "  select id,name from currency  order by name  "
        Else
            My_SQL = "  select id,code from currency  order by code  "
        End If

        fill_combo DcCurrency, My_SQL

    End If

End Sub

Private Sub DcEmp_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
        
        My_SQL = "  select Emp_ID,Emp_name  from TblEmployee  "
        fill_combo DcEmp, My_SQL
    End If

End Sub

Private Sub DCEmp1_Change()
Dim fullcode As String
Dim DefaultSalesPersonId As Integer
  
        TxtCustCode2.text = GetSalespersonDetail(val(DcEmp1.BoundText))
        
End Sub

Private Sub DCEmp1_Click(Area As Integer)
DCEmp1_Change
End Sub

Private Sub DCPreFix_KeyUp(KeyCode As Integer, _
                           Shift As Integer)
 
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetPrefix Me.DCPreFix, 0, 0 'val(branch_id)

    End If
        
End Sub

Private Sub employee_details_Click()
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    VSFlexGrid1.Enabled = True
          
    If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
        Frame10.Visible = True
        Frame10.Enabled = True

        current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode"))

        If SystemOptions.UserInterface = ArabicInterface Then
            Frame10.Caption = "«·⁄„«·Â ›Ì  «·⁄„·Ì… —ﬁ„ : " & current_opr
        Else
            Frame10.Caption = "Labors for Operation NO : " & current_opr
        End If
        
        'Me.txt_emp_salary.text = IIf(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("salary")) = "", 0, VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("salary")))
        'Me.txt_employee_count.text = IIf(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("count")) = "", 0, VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("count")))
        
        Frame10.Visible = True
        StrSQL = "SELECT  * FROM  opr_employee_details Where   opr_fullcode='" & current_opr & "'"

        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            RsDev.MoveFirst
    
            With Me.VSFlexGrid1
                .Rows = .FixedRows + RsDev.RecordCount

                For i = .FixedRows To .Rows - 1
 
                    .TextMatrix(i, .ColIndex("LineNo")) = i
            
                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("Emp_ID").value), "", RsDev("Emp_ID").value)
            
                    .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(RsDev("Emp_Code").value), "", RsDev("Emp_Code").value)
            
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
            
                    .TextMatrix(i, .ColIndex("jobid")) = IIf(IsNull(RsDev("JobTypeID").value), "", RsDev("JobTypeID").value)
                    .TextMatrix(i, .ColIndex("jobname")) = IIf(IsNull(RsDev("JobTypeName").value), "", RsDev("JobTypeName").value)
                    .TextMatrix(i, .ColIndex("daysalary")) = IIf(IsNull(RsDev("daysalary").value), 0, RsDev("daysalary").value)
                    .TextMatrix(i, .ColIndex("count")) = IIf(IsNull(RsDev("count").value), 0, RsDev("count").value)
        
                    .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev("total").value), 0, RsDev("total").value)

                    RsDev.MoveNext
                Next i

            End With
    
        End If
    
    End If

    calcnets
    ReLineGrid
End Sub

Private Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                 ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    'Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg_Journal

        Select Case .ColKey(Col)
 
            Case "by"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("sub_contractor_id"), False, True)
                .TextMatrix(Row, .ColIndex("sub_contractor_id")) = StrAccountCode
                '.TextMatrix(Row, .ColIndex("id")) = get_Expenses_id(StrAccountCode)
         
        End Select
   
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid

End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "by"
                Exit Sub
        End Select

    End With

    Fg_Journal.ComboList = ""

End Sub

Private Sub Fg_Journal_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    With Me.Fg_Journal

        Select Case .ColKey(Col)

                 Case "opera"
                  LngRow = Row

 LngCol = Col
             ' ItemProductionDate Row, Col, , 1
                Load FrmProcessOfProject
                FrmProcessOfProject.show vbModal

                    
                End Select
                End With
End Sub

Private Sub Fg_Journal_Click()

    If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("fullcode")) = "" Then
        
        current_terms = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("fullcode"))
        ProjectDes_ID = val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("oprid")))
    End If

End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)
        Case "opera"
         .ColComboList(.ColIndex("opera")) = "..."
        

            Case "by"
                StrSQL = "select * from TblCustemers where Type=3"
  Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rs.RecordCount = 0 Then Exit Sub
              
                StrComboList = Fg_Journal.BuildComboList(rs, "CusName", "CusID")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set SaleReport = Nothing
End Sub

Private Sub GridSub_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim rat As Integer
With GridSub
Select Case .ColKey(Col)
Case "rate"
rat = IIf(Not IsNumeric(.TextMatrix(Row, .ColIndex("rate"))), 0, .TextMatrix(Row, .ColIndex("rate")))
  .TextMatrix(Row, .ColIndex("SubValue")) = rat / 100 * val(total_after_discount.text)
    
        End Select
            If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
ReLineGrid2
 
    End With
End Sub


Private Sub RemoveGridRow()

    With Me.GridSub

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid2
End Sub


Private Sub GridSub_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim LngItemID As Long
    Dim LngStoreID As Long
    Dim rdate As Date
  ' Dim frm As FrmGridAddItemComment
    Dim Frm1 As FrmRegesterDateProject

    'On Error GoTo ErrTrap

    With Me.GridSub

        Select Case .ColKey(Col)

                 Case "subdate"
                  LngRow = Row

 LngCol = Col
             ' ItemProductionDate Row, Col, , 1
                Load FrmRegesterDateProject
                FrmRegesterDateProject.show

                    
                End Select
                End With

End Sub

Private Sub GridSub_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With Me.GridSub

        Select Case .ColKey(Col)

                 Case "subdate"
    
            .ColComboList(.ColIndex("subdate")) = "..."
            Case "rate"
          If .ColComboList(.ColIndex("subdate")) = "" Then
          MsgBox "ÌÃ» «Œ Ì«—  «—ÌŒ «·œ›⁄Â «Ê·«"
          Exit Sub
          End If
              Case "SubValue"
          If .ColComboList(.ColIndex("subdate")) = "" Then
          MsgBox "ÌÃ» «Œ Ì«—  «—ÌŒ «·œ›⁄Â «Ê·«"
          Exit Sub
          End If
            End Select
            End With
End Sub

Private Sub Label37_Click()
Fra(2).Visible = False
End Sub

Private Sub Label40_Click()
    Fra(3).Visible = False
End Sub

Private Sub OptType_Click(Index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.text)
End Sub

Private Sub OptType1_Click(Index As Integer)
    Me.TxtOpenBalance1.Enabled = Not OptType1(2).value
    Me.TxtOpenBalance1.text = IIf(OptType1(2).value = True, 0, Me.TxtOpenBalance1.text)
End Sub

Private Sub OptType2_Click(Index As Integer)
    Me.TxtOpenBalance2.Enabled = Not OptType2(2).value
    Me.TxtOpenBalance2.text = IIf(OptType2(2).value = True, 0, Me.TxtOpenBalance2.text)
End Sub

Private Sub OptType3_Click(Index As Integer)
    Me.TxtOpenBalance3.Enabled = Not OptType3(2).value
    Me.TxtOpenBalance3.text = IIf(OptType3(2).value = True, 0, Me.TxtOpenBalance3.text)
End Sub

Private Sub OptType4_Click(Index As Integer)
    Me.TxtOpenBalance4.Enabled = Not OptType4(2).value
    Me.TxtOpenBalance4.text = IIf(OptType4(2).value = True, 0, Me.TxtOpenBalance4.text)
End Sub

Private Sub Form_Load()
'equepment
    Dim My_SQL As String

    My_SQL = "  select Account_Code,CusID from TblCustemers  where type=1 "

    fill_combo DcAccount1, My_SQL

    My_SQL = "  select Account_Code,CusID from TblCustemers  where type=1  "

    fill_combo DcAccount3, My_SQL
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetPrefix Me.DCPreFix, 0, 0 'val(branch_id)
Dcombos.GetSalesRepData Me.DcEmp1

Dcombos.GetSection Me.DcbDept




    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  where type=1  "
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  where type=1  "
    End If

    fill_combo DcAccount2, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select id,code from currency  order by name  "
    Else
        My_SQL = "  select id,code from currency  order by code  "
    End If

    fill_combo DcCurrency, My_SQL
 
    My_SQL = "  select JobTypeID,JobTypeName from TblEmpJobsTypes  order by JobTypeName  "
    fill_combo dcJobTypeName, My_SQL

    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  where type=1  "
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  where type=1  "
    End If

    fill_combo DcAccount4, My_SQL

    My_SQL = "  select id,name from project_status  "
    fill_combo DataCombo1, My_SQL
    
If SystemOptions.UserInterface = ArabicInterface Then
My_SQL = "  select id,name from contract_type  "
Else
My_SQL = "  select id,namee from contract_type  "
End If
    fill_combo DataCombo5, My_SQL

    'If SystemOptions.UserInterface = ArabicInterface Then
    'My_SQL = "  select branch_id,branch_name from branches  "
    ' Else
    ' My_SQL = "  select branch_id,branch_namee from branches  "
    ' End If
    '
    'fill_combo Dcbranch, My_SQL

    'Dim Dcombos As ClsDataCombos
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.Dcbranch.Enabled = False
    End If

    Dcombos.GetBranches Dcbranch

    My_SQL = "  select Emp_ID,Emp_name  from TblEmployee  "

    fill_combo DcEmp, My_SQL

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    '    Exit Sub
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Set NewGrid.Grid = FG
    NewGrid.GridTrans = INVENTORYIN
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.TxtTotal = XPTxtSum
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DcboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.DtpBillDate = Me.XPDtbBill

    Set NewGrid.LblTotalQty = Me.LblTotalQty
    NewGrid.fillGrid
    'FG.WallPaper = BGround.Picture
 
    'SetDtpickerDate XPDtbBill
    'Set Dcombos = New ClsDataCombos
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'LoadSettings
    Set rs = New ADODB.Recordset
    StrSQL = "select * From projects "
      If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
        StrSQL = StrSQL & " where   branch_no=" & Current_branch
    End If
    
     StrSQL = StrSQL & " order by id  "
     
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
   ' rs.Open "projects", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"
    'Exit Sub
  
    If SystemOptions.UserInterface = EnglishInterface Then
        ChangeLang
    End If

    ' Me.Width = 10000
  
    If OPEN_NEW_SCREEN = True Then
        Command1_Click (0)
    End If

End Sub

Function ChangeLang()
Option6.Caption = "Estimation"
Option7.Caption = "Actual"
Frame8.Caption = "Type"
Ptype(0).Caption = "New"
Ptype(1).Caption = "Opening"
Label42.Caption = "Disc.%"
Command1(8).Caption = "Payments Data"

    Fra(7).Caption = "Accountant Data"
    Fra(8).Caption = "Opening Expenses  Balances "
    OptType(0).Caption = "Depit"
    OptType(1).Caption = "Credit"
    OptType(2).Caption = "Na"
    Lbl(14).Caption = "Balance"
    Lbl(13).Caption = "Date"

    Fra(9).Caption = "Opening Revenues Balances "
    OptType1(0).Caption = "Depit"
    OptType1(1).Caption = "Credit"
    OptType1(2).Caption = "Na"
    Lbl(15).Caption = "Balance"
    Lbl(16).Caption = "Date"

    Fra(10).Caption = "Opening  Materials Balances "
    OptType2(0).Caption = "Depit"
    OptType2(1).Caption = "Credit"
    OptType2(2).Caption = "Na"
    Lbl(18).Caption = "Balance"
    Lbl(17).Caption = "Date"

    Fra(1).Caption = "Opening  Salaries Balances "
    OptType3(0).Caption = "Depit"
    OptType3(1).Caption = "Credit"
    OptType3(2).Caption = "Na"
    Lbl(10).Caption = "Balance"
    Lbl(11).Caption = "Date"

    Fra(0).Caption = "Opening Invoices Balances  "
    OptType4(0).Caption = "Depit"
    OptType4(1).Caption = "Credit"
    OptType4(2).Caption = "Na"
    Lbl(8).Caption = "Balance"
    Lbl(9).Caption = "Date"

    Command1(4).Caption = "Opening Balances"

    Label14.Caption = " Start D."
    Command3.Caption = "Reports"
    Label26.Caption = "Branch"
    temp = XPBtnMove(1).left
    XPBtnMove(1).left = XPBtnMove(2).left
    XPBtnMove(2).left = temp
    Label36.Caption = "nearest end"
    Label35.Caption = "Manger"
    temp = XPBtnMove(0).left
    XPBtnMove(0).left = XPBtnMove(3).left
    XPBtnMove(3).left = temp
    SetInterface Me
    Label16.Caption = "End User ID"
    Label15.Caption = "End User Name"
    Label23.Caption = "Sub-contractor"
    Label24.Caption = "contraactor Name"
    Label38.Caption = "Name Eng"
    Label6.Caption = "Project Code"
    Label5.Caption = "Staus"
    Label1.Caption = "Project Name"
    Label8.Caption = "Con. Type"
    Label20.Caption = "Project Cost"
    Label21.Caption = "Duration"
    Label17.Caption = "End D."
    Label18.Caption = "Earliest end."
    Label22.Caption = "Notes"
    Frame4.Caption = "Color Map"
    Label34.Caption = "Critical"
    'Label21.Caption = "Expanses Account"
    'Label22.Caption = "Revenue Account"
    'Label17.Caption = "Items"
    'Label18.Caption = "Item Description"
    Label19.Caption = "Currency"

    Frame5.Caption = "Terms Data"
    terms_operations(0).Caption = "Terms Operations"
    Frame12.Caption = "Expenses"
    opr_Expenses(1).Caption = "Return To Opr."
    Lbl(6).Caption = "Total Expenses"

    With Me.VSFlexGrid3
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = "Expenses Names"
        .TextMatrix(0, .ColIndex("value")) = "value"
 .TextMatrix(0, .ColIndex("EsToal")) = "Estimated Value"
 
        .TextMatrix(0, .ColIndex("des")) = "des"
 
    End With

    Frame1.Caption = "Items"

    '   txtid.Alignment = 0
    DataCombo1.RightToLeft = False
    '
    CMD_language.Caption = "⁄—»Ì"
    '  Frame4.Visible = True
    Frame3.Visible = True
    '    Frame8.Visible = True
    
    Label9.Caption = "Projects Data"
    Me.Caption = Label9.Caption
  
    Command1(0).Caption = "new"
    Command1(1).Caption = "save"
    Command1(2).Caption = "Attachments"
    '  SuperLabel2.text = "Search"
    '  Command1(4).Caption = "By ID"
    Command1(5).Caption = "Search"
   
    Label32.Caption = "Discount"
    Label33.Caption = "Net Cost"
    Label31.Caption = "Total"
    Command1(3).Caption = "Edit"
    CMDViewGantt(2).Caption = "View Gantt "
    Label12.Caption = "Period"
Label41.Caption = "Employee"
opr_Expenses(2).Caption = "Equipments"

    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("fullcode")) = "Term Code"

        .TextMatrix(0, .ColIndex("des")) = "Des"
        .TextMatrix(0, .ColIndex("qty")) = "Qty"
        .TextMatrix(0, .ColIndex("cost")) = "Cost"
        .TextMatrix(0, .ColIndex("total")) = "Total"
        
         .TextMatrix(0, .ColIndex("esQty")) = "Estm. Qty"
        .TextMatrix(0, .ColIndex("EsPrice")) = "Estm.  Cost"
        .TextMatrix(0, .ColIndex("EstTotal")) = "Estm.  Total"
     
        .TextMatrix(0, .ColIndex("discount")) = "discount"
        .TextMatrix(0, .ColIndex("net")) = "net"
        .TextMatrix(0, .ColIndex("By")) = "Sub-contarctor"
    End With
    Cmd.Caption = "Delete"
Frame6.Caption = "Equpiments Data"
Lbl(7).Caption = "Totals"
opr_Expenses(3).Caption = "Return To Operations"
Fra(2).Caption = "Payments Data"
Lbl(19).Caption = "Estim. Qty"
Lbl(12).Caption = "Estim. Price"
  With Me.GridSub
  
 



        .TextMatrix(0, .ColIndex("id")) = "id"
        .TextMatrix(0, .ColIndex("subdate")) = "date"

        .TextMatrix(0, .ColIndex("DesTerm")) = "Finished Term/Process"
        .TextMatrix(0, .ColIndex("rate")) = "Rate"
        .TextMatrix(0, .ColIndex("SubValue")) = "Value"
         
         .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
     
    End With
    
  With Me.VSFlexGrid4
  


        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("FixedAsset")) = "Equipment"

        .TextMatrix(0, .ColIndex("EstHour")) = "Estimated Hour"
        .TextMatrix(0, .ColIndex("ActualHour")) = "Actual Hour"
        .TextMatrix(0, .ColIndex("TotalEs")) = "Total Estimated"
        .TextMatrix(0, .ColIndex("value")) = "Actual Total "
         .TextMatrix(0, .ColIndex("des")) = "Des"
     
    End With
    
    Frame11.Caption = "Terms Operations"

    With Me.VSFlexGrid2
 
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("Symbol")) = "Symbol"
        .TextMatrix(0, .ColIndex("Pre")) = "Based On"
        .TextMatrix(0, .ColIndex("Earlystartweek")) = "E. start " & getoprTitle
        .TextMatrix(0, .ColIndex("startweek")) = "start " & getoprTitle
        .TextMatrix(0, .ColIndex("EarlyEndWeek")) = "E. End " & getoprTitle
        .TextMatrix(0, .ColIndex("EndWeek")) = "End " & getoprTitle
        .TextMatrix(0, .ColIndex("Critical")) = "Critical"

        .TextMatrix(0, .ColIndex("fullcode")) = "OPR Code"
        .TextMatrix(0, .ColIndex("name")) = "OPR Name"
        .TextMatrix(0, .ColIndex("period")) = "p " & getoprTitle
        .TextMatrix(0, .ColIndex("period1")) = "Slack " & getoprTitle

        .TextMatrix(0, .ColIndex("total_items")) = "total items Cost"
        .TextMatrix(0, .ColIndex("total_salary")) = "total salary"
        .TextMatrix(0, .ColIndex("total_expenses")) = "total expenses"
        .TextMatrix(0, .ColIndex("total")) = "Total"
        .TextMatrix(0, .ColIndex("qty")) = "Qty"
        .TextMatrix(0, .ColIndex("unitname")) = "Unit Name"
        .TextMatrix(0, .ColIndex("periodView")) = "Default Period"
        .TextMatrix(0, .ColIndex("Actperiod")) = "Actual Period"
    End With

    opr_items(0).Caption = " Items"
    employee_details.Caption = " Labors "
    opr_Expenses(0).Caption = " Expenses"
    terms_operations(1).Caption = "Return To  Terms"
    Label28.Caption = "Total"
    Command1(6).Caption = "Undo"
    Command1(7).Caption = "Delete"
    Lbl(4).Visible = False
    Lbl(5).Visible = False
    Shape1.Visible = False
    Frame10.Caption = "Labors  Data"
    Label27.Caption = "No of Labors "
    Label29.Caption = "Total salaaries"
    opr_emplyees_name.Caption = "Return To Opr."

    Label30.Caption = "ID"
    Label4.Caption = "Count"
    Label11.Caption = "W. Days"
    Label10.Caption = "Start Date"
    Command2.Caption = "Add"

    Label3.Caption = "Select Job Type"
    Option4.Caption = "Estimation"
    Option5.Caption = "Allocation"

    With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("code")) = "Code"
        .TextMatrix(0, .ColIndex("name")) = "Name"
        .TextMatrix(0, .ColIndex("jobname")) = "Job Name"
        .TextMatrix(0, .ColIndex("daysalary")) = "Day Salary"
        .TextMatrix(0, .ColIndex("Count")) = "No.Of.Days"
        .TextMatrix(0, .ColIndex("total")) = "Total"

        .TextMatrix(0, .ColIndex("des")) = "Remark"
 
    End With

    Frame1.Caption = "OPR Items"
    Lbl(31).Caption = "Item Code"
    Lbl(30).Caption = "Item Name"
    Lbl(29).Caption = "Status"
    Lbl(28).Caption = "Serial"
    Lbl(27).Caption = "QTY"
    Lbl(26).Caption = "Price"
    Lbl(0).Caption = "Avilable"
    Lbl(1).Caption = "Reserved"
    Lbl(3).Caption = "ON order"
    Lbl(2).Caption = "Total"
    opr_items(1).Caption = "Return To Opr."
 
End Function

Function SaveAutoVoucher(Vtype As Integer)
    'Vtype = 0   Œ’Ì’ ›⁄·Ï
    ''Vtype = 3  Œ’Ì’  ﬁœÌ—Ì ›ﬁÿ
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim LngDevID As Long
    Dim voucherid As Integer
    'On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
        
    rs.Open "opr_Employee", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    If Me.TxtModFlg.text <> "R" Then
 
        Cn.Execute "delete opr_Employee where auto=1 and Project_id=" & val(Me.txt_project_id.text)
    
        rs.AddNew
        voucherid = CStr(new_id("opr_Employee", "ID", "", True))
        rs("ID").value = voucherid
   
        rs("Start_date").value = XPDtbTrans.value
        rs("Project_id").value = val(txt_project_id)
        rs("opr_type").value = Vtype
        'Vtype = 0   Œ’Ì’ ›⁄·Ï
        ''Vtype = 3  Œ’Ì’  ﬁœÌ—Ì ›ﬁÿ
        rs("Auto").value = 1
        rs("recorddate").value = Date
        rs("term_Fullcode").value = current_terms
     
        rs("opr_Fullcode").value = current_opr

        rs.update
    
        Set RsDev = New ADODB.Recordset
        
        RsDev.Open "opr_employee_details", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
        Dim i As Integer

        With Me.VSFlexGrid1

            For i = .FixedRows To .Rows - 1

                If .TextMatrix(i, .ColIndex("id")) <> "" Then
         
                    RsDev.AddNew
                    RsDev("pk_id").value = voucherid
                    RsDev("emp_code").value = .TextMatrix(i, .ColIndex("code"))
                    RsDev("emp_name").value = .TextMatrix(i, .ColIndex("name"))
                    RsDev("JobTypeName").value = .TextMatrix(i, .ColIndex("jobname"))
                    RsDev("JobTypeID").value = .TextMatrix(i, .ColIndex("jobid"))
            
                    RsDev("Emp_id").value = .TextMatrix(i, .ColIndex("id"))
                    RsDev("Start_date").value = XPDtbTrans.value
                    RsDev("Project_id").value = val(Me.txt_project_id)
                    RsDev("opr_type").value = Vtype
            
                    RsDev("term_Fullcode").value = current_terms
           
                    RsDev("opr_Fullcode").value = current_opr
                    RsDev("daysalary").value = val(.TextMatrix(i, .ColIndex("daysalary")))
                    RsDev("count").value = val(.TextMatrix(i, .ColIndex("count")))
                    RsDev("total").value = val(.TextMatrix(i, .ColIndex("total")))
     
                    If Vtype = 0 Then
                        save_employee_current_status val(Me.txt_project_id), current_terms, current_opr, val(.TextMatrix(i, .ColIndex("id")))
                    End If

                    RsDev.update
                    
                End If
            
                '
            Next i

        End With
 
    End If

    Exit Function
ErrTrap:
     
End Function

Private Sub opr_emplyees_name_Click()
    calcnets
    VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("total_salary")) = val(txt_emp_salary)
    ReLineGrid
    Frame10.Visible = False
    Set RsDev = New ADODB.Recordset
    Dim sql As String

    '«‰‘«¡ ”‰œ  Œ’Ì’ «·Ì » «—ÌŒ «·„Õœœ
    If Option4.value = True Then
        SaveAutoVoucher (3) ' ﬁœÌ—
    ElseIf Option5.value = True Then
        SaveAutoVoucher (0) '›⁄·Ì
    End If
 
    Exit Sub
 
    If Option5.value = True Then

        With VSFlexGrid1

            For i = .FixedRows To .Rows - 2

                If .TextMatrix(i, .ColIndex("id")) <> "" Then
                    sql = "Select * from TblEmployee where Emp_ID=" & .TextMatrix(i, .ColIndex("id"))
                    RsDev.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
         
                    RsDev("opr_fullcode").value = current_opr
                    RsDev("project_id").value = val(Me.txt_project_id)
                    RsDev("term_id").value = val(current_terms)
                    RsDev("opr_id").value = val(current_opr)
        
                    RsDev.update
                    RsDev.Close
                End If

            Next i
    
        End With

    End If

    'VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("salary")) = IIf(Not IsNumeric(Me.txt_emp_salary.text), 0, Me.txt_emp_salary.text)
    'VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("count")) = IIf(Not IsNumeric(Me.txt_employee_count.text), 0, Me.txt_employee_count.text)
 
    sql = "Select * from terms_operations where fullcode='" & current_opr & "'"
    RsDev.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsDev.RecordCount <= 0 Then Exit Sub
    RsDev("count").value = IIf(Not IsNumeric(Me.txt_employee_count.text), 0, Me.txt_employee_count.text)
        
    RsDev("salary").value = IIf(Not IsNumeric(Me.txt_emp_salary.text), 0, Me.txt_emp_salary.text)
        
    RsDev.update
    RsDev.Close

End Sub

Private Sub opr_expenses_Click(Index As Integer)

    Select Case Index
Case 2
Frame6.Visible = True

Case 3
Frame6.Visible = False

        Case 0
  
            VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid3.Rows = 2
            VSFlexGrid3.Enabled = True

            If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
                Frame12.Visible = True

                current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode"))
                Retrive3 current_opr

                If SystemOptions.UserInterface = ArabicInterface Then
                    Frame1.Caption = "„’«—Ì› «·⁄„·Ì… —ﬁ„ :   " & "  " & current_opr
                Else
                    Frame1.Caption = "Expenses For Operation NO: " & "  " & current_opr
                End If
        
                XPTxtSum.text = 0
            End If

        Case 1

            VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("total_expenses")) = val(txt_expenses_total)
            ReLineGrid

            StrSQL = "Delete From opr_expenses where opr_fullcode='" & current_opr & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            Set RSTransDetails = New ADODB.Recordset
            RSTransDetails.Open "[opr_expenses]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

            For RowNum = 1 To VSFlexGrid3.Rows - 1

                If VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("ExpensesID")) <> "" Then
  
                    RSTransDetails.AddNew
    
                    RSTransDetails("opr_fullcode").value = current_opr
     
                    RSTransDetails("ExpensesID").value = IIf((VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("ExpensesID")) = ""), Null, val(VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("ExpensesID"))))
                    RSTransDetails("AccountCode").value = IIf((VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("AccountCode")) = ""), Null, VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("AccountCode")))
                    RSTransDetails("AccountName").value = IIf((VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("AccountName")) = ""), Null, VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("AccountName")))
                    RSTransDetails("value").value = IIf((VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("value")) = ""), Null, val(VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("value"))))
                    RSTransDetails("des").value = IIf((VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("des")) = ""), Null, VSFlexGrid3.TextMatrix(RowNum, VSFlexGrid3.ColIndex("des")))
 
                    RSTransDetails.update
                End If

            Next

            Frame12.Visible = False

    End Select

End Sub

Private Sub opr_items_Click(Index As Integer)
    Dim currentqty As Double

    Select Case Index

        Case 0

            If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
                Frame1.Visible = True
                currentqty = val(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("qty")))
                current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("OPRIDD"))
                Retrive2 current_opr, currentqty

                If SystemOptions.UserInterface = ArabicInterface Then
                    Frame1.Caption = "„Ê«œ «·⁄„·Ì… —ﬁ„ :   " & "  " & current_opr
                Else
                    Frame1.Caption = "Items For Operations NO :   " & "  " & current_opr
                End If
        
                ' XPTxtSum.text = 0
                With FG
                    Me.XPTxtSum.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Valu"), .Rows - 1, .ColIndex("Valu"))
                End With

            End If

        Case 1
            VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("total_items")) = val(XPTxtSum)
            ReLineGrid
            StrSQL = "Delete From Transaction_Details where  (payed is null )  and  opr_fullcode='" & current_opr & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            Set RSTransDetails = New ADODB.Recordset
            RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

            For RowNum = 1 To FG.Rows - 1

                If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
  
                    RSTransDetails.AddNew
    
                    RSTransDetails("opr_fullcode").value = current_opr
                    RSTransDetails("Project_id").value = val(txt_project_id.text)
                    RSTransDetails("term_id").value = val(current_terms)
                    RSTransDetails("opr_id").value = val(current_opr)
    
                    RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                    RSTransDetails("Price").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                    RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                    RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
 
                    RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
 
                    RSTransDetails.update
                End If

            Next

            Frame1.Visible = False

Case 2
Frame6.Enabled = True

Case 3
Frame6.Enabled = False

    End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

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
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Option4_Click()
    XPDtbTrans.Enabled = False
End Sub

Private Sub Option5_Click()
    XPDtbTrans.Enabled = True
End Sub

Private Sub terms_operations_Click(Index As Integer)
Dim RsDetails12 As ADODB.Recordset
Dim RsDetails1 As ADODB.Recordset
Dim RsDetails11 As ADODB.Recordset
  Dim astrSplit2tems2() As String
Dim astrSplitItems() As String
Dim J As Integer
  Dim st As String
    Dim nElements As Integer
    Select Case Index

        Case 0

            If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("fullcode")) = "" Then
                Frame11.Visible = True
        
                current_terms = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("fullcode"))
                ProjectDes_ID = val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("oprid")))
                retrive1 current_terms

                If SystemOptions.UserInterface = ArabicInterface Then
                    Frame11.Caption = "⁄„·Ì«  «·»‰œ —ﬁ„ : " & current_terms
                Else
                    Frame11.Caption = "Operations For Term No: " & current_terms
                End If
            End If

        Case 1
            ReLineGrid current_terms

          'ds
        
            ' ⁄„·Ì«  «·»‰Êœ
            '"«·„Ê«œœ"
            'Dim StrSQL As String
                                   Set RsDetails12 = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblExpensiveOper Where (1 = -1)"
   RsDetails12.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
                               Set RsDetails11 = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblEquepment Where (1 = -1)"
   RsDetails11.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
                        Set RsDetails1 = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblEmpOper Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
              Set RsDetails = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblMatrials Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
            Set RsDev = New ADODB.Recordset
            RsDev.Open "terms_operations", Cn, adOpenStatic, adLockOptimistic, adCmdTable
Dim OPR As Integer
            Dim i As Integer

            With Me.VSFlexGrid2

                For i = .FixedRows To .Rows - 1

                    '
                    If .TextMatrix(i, .ColIndex("name")) <> "" Then

                        RsDev.AddNew

                        If .TextMatrix(i, .ColIndex("fullcode")) = "" Then
                            RsDev("fullcode").value = current_terms & "-" & .TextMatrix(i, .ColIndex("LineNo"))
                        Else
                            RsDev("fullcode").value = .TextMatrix(i, .ColIndex("fullcode"))
                        End If
                     StrSQL = "Delete From terms_operations Where fullcode ='" & .TextMatrix(i, .ColIndex("fullcode")) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecor
                 StrSQL = "Delete From TblMatrials Where OperCode ='" & .TextMatrix(i, .ColIndex("fullcode")) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
             StrSQL = "Delete From TblEmpOper Where OperCode ='" & .TextMatrix(i, .ColIndex("fullcode")) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
             StrSQL = "Delete From TblEquepment Where OperCode ='" & .TextMatrix(i, .ColIndex("fullcode")) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
              StrSQL = "Delete From TblExpensiveOper Where OperCode ='" & .TextMatrix(i, .ColIndex("fullcode")) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
    '    If Me.TxtModFlg.text = "E" Then
       OPR = val(.TextMatrix(i, .ColIndex("id")))
       If Me.Checked(0, OPR) = True Then
       Else
       OPR = 1
       maxx 0, OPR
       End If
      ' If Me.TxtModFlg.text = "N" Then
     '   OPR = 1
     '  maxx 0, OPR
      ' End If
       
           '   End If
                        RsDev("project_id").value = val(Me.txt_project_id.text)
                        RsDev("term_fullcode").value = current_terms
                        RsDev("ProjectDes_ID").value = ProjectDes_ID
                        RsDev("id").value = OPR
                       ' RsDev("id").value = .TextMatrix(i, .ColIndex("LineNo"))
                        RsDev("total").value = IIf(Not IsNumeric(.TextMatrix(i, .ColIndex("total"))), 0, .TextMatrix(i, .ColIndex("total")))
                        RsDev("name").value = .TextMatrix(i, .ColIndex("name"))
                        RsDev("period").value = IIf(.TextMatrix(i, .ColIndex("period")) = "", 0, .TextMatrix(i, .ColIndex("period")))
                        RsDev("count").value = IIf(.TextMatrix(i, .ColIndex("count")) = "", 0, .TextMatrix(i, .ColIndex("count")))
                        RsDev("salary").value = IIf(.TextMatrix(i, .ColIndex("salary")) = "", 0, .TextMatrix(i, .ColIndex("salary")))
                        RsDev("total_items").value = IIf(.TextMatrix(i, .ColIndex("total_items")) = "", 0, .TextMatrix(i, .ColIndex("total_items")))
                        RsDev("total_salary").value = IIf(.TextMatrix(i, .ColIndex("total_salary")) = "", 0, .TextMatrix(i, .ColIndex("total_salary")))
                        RsDev("total_expenses").value = IIf(.TextMatrix(i, .ColIndex("total_expenses")) = "", 0, .TextMatrix(i, .ColIndex("total_expenses")))
''//
                      RsDev("expen").value = IIf(.TextMatrix(i, .ColIndex("expen")) = "", "", .TextMatrix(i, .ColIndex("expen")))
                       RsDev("eq").value = IIf(.TextMatrix(i, .ColIndex("eq")) = "", "", .TextMatrix(i, .ColIndex("eq")))
                      RsDev("emps").value = IIf(.TextMatrix(i, .ColIndex("emps")) = "", "", .TextMatrix(i, .ColIndex("emps")))
                       RsDev("matrials").value = IIf(.TextMatrix(i, .ColIndex("matrials")) = "", "", .TextMatrix(i, .ColIndex("matrials")))
''//
                        RsDev("Symbol").value = IIf(.TextMatrix(i, .ColIndex("Symbol")) = "", "", .TextMatrix(i, .ColIndex("Symbol")))
                        RsDev("Pre").value = IIf(.TextMatrix(i, .ColIndex("Pre")) = "", "", .TextMatrix(i, .ColIndex("Pre")))
                        RsDev("period1").value = IIf(.TextMatrix(i, .ColIndex("period1")) = "", 0, .TextMatrix(i, .ColIndex("period1")))
                        RsDev("Earlystartweek").value = IIf(.TextMatrix(i, .ColIndex("Earlystartweek")) = "", 0, .TextMatrix(i, .ColIndex("Earlystartweek")))
                        RsDev("startweek").value = IIf(.TextMatrix(i, .ColIndex("startweek")) = "", 0, .TextMatrix(i, .ColIndex("startweek")))
                        RsDev("EarlyEndWeek").value = IIf(.TextMatrix(i, .ColIndex("EarlyEndWeek")) = "", 0, .TextMatrix(i, .ColIndex("EarlyEndWeek")))
                        RsDev("EndWeek").value = IIf(.TextMatrix(i, .ColIndex("EndWeek")) = "", 0, .TextMatrix(i, .ColIndex("EndWeek")))
                        RsDev("Critical").value = IIf(.TextMatrix(i, .ColIndex("Critical")) = "", 0, .TextMatrix(i, .ColIndex("Critical")))
                        RsDev("OPRIDD").value = IIf(.TextMatrix(i, .ColIndex("OPRIDD")) = "", 0, .TextMatrix(i, .ColIndex("OPRIDD")))
                        RsDev("Actperiod").value = IIf(.TextMatrix(i, .ColIndex("Actperiod")) = "", 0, .TextMatrix(i, .ColIndex("Actperiod")))
                        RsDev("periodView").value = IIf(.TextMatrix(i, .ColIndex("periodView")) = "", "", .TextMatrix(i, .ColIndex("periodView")))
                        RsDev("qty").value = IIf(.TextMatrix(i, .ColIndex("qty")) = "", 0, .TextMatrix(i, .ColIndex("qty")))
                        RsDev("unitname").value = IIf(.TextMatrix(i, .ColIndex("unitname")) = "", 0, .TextMatrix(i, .ColIndex("unitname")))
                        RsDev("unitid").value = IIf(.TextMatrix(i, .ColIndex("unitid")) = "", 0, .TextMatrix(i, .ColIndex("unitid")))
                        RsDev.update
                    ''///// «·„Ê«œ
                                   If VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("matrials")) <> "" Then
          st = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("matrials"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For J = 0 To nElements - 1
          RsDetails.AddNew
                   astrSplit2tems2 = Split(astrSplitItems(J), "#")
        RsDetails("OperCode").value = .TextMatrix(i, .ColIndex("fullcode"))
         RsDetails("ProjectID").value = val(Me.txt_project_id.text)
         RsDetails("Pand").value = ProjectDes_ID
         RsDetails("Opr").value = IIf(IsNull(RsDev("id").value), Null, RsDev("id").value)
         RsDetails("ItemID").value = val(astrSplit2tems2(0))
         RsDetails("Count").value = val(astrSplit2tems2(1))
         RsDetails("Price").value = val(astrSplit2tems2(2))
         RsDetails("Quntapro").value = val(astrSplit2tems2(3))
         RsDetails("priceapro").value = val(astrSplit2tems2(4))
                
         RsDetails.update
         Next J
          End If
      ''//////////////
          
          End If
    
                            ''///// «·⁄„«·Â
                                   If VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("emps")) <> "" Then
          st = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("emps"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For J = 0 To nElements - 1
          RsDetails1.AddNew
         astrSplit2tems2 = Split(astrSplitItems(J), "#")
         RsDetails1("OperCode").value = .TextMatrix(i, .ColIndex("fullcode"))
         RsDetails1("ProjectID").value = val(Me.txt_project_id.text)
         RsDetails1("Pand").value = ProjectDes_ID
         RsDetails1("Opr").value = IIf(IsNull(RsDev("id").value), Null, RsDev("id").value)
         RsDetails1("EmpID").value = val(astrSplit2tems2(0))
         RsDetails1("JobID").value = val(astrSplit2tems2(1))
         RsDetails1("daysalary").value = val(astrSplit2tems2(2))
         RsDetails1("Count").value = val(astrSplit2tems2(3))
                         
         RsDetails1.update
         Next J
          End If
      ''//////////////
                             ''///// «·„⁄œ« 
                                   If VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("eq")) <> "" Then
          st = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("eq"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For J = 0 To nElements - 1
          RsDetails11.AddNew
         astrSplit2tems2 = Split(astrSplitItems(J), "#")
         RsDetails11("OperCode").value = .TextMatrix(i, .ColIndex("fullcode"))
         RsDetails11("ProjectID").value = val(Me.txt_project_id.text)
         RsDetails11("Pand").value = ProjectDes_ID
         RsDetails11("Opr").value = IIf(IsNull(RsDev("id").value), Null, RsDev("id").value)
         RsDetails11("ExpensesID").value = val(astrSplit2tems2(0))
         RsDetails11("EstHour").value = val(astrSplit2tems2(1))
         RsDetails11("ActualHour").value = val(astrSplit2tems2(2))
         RsDetails11("TotalEs").value = val(astrSplit2tems2(3))
         RsDetails11("value").value = val(astrSplit2tems2(4))
         RsDetails11("des").value = astrSplit2tems2(5)
                         
         RsDetails11.update
         Next J
          End If
      ''//////////////
                                   ''///// «·„’«—Ì›
         If VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("expen")) <> "" Then
          st = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("expen"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
         For J = 0 To nElements - 1
          RsDetails12.AddNew
         astrSplit2tems2 = Split(astrSplitItems(J), "#")
         RsDetails12("OperCode").value = .TextMatrix(i, .ColIndex("fullcode"))
         RsDetails12("ProjectID").value = val(Me.txt_project_id.text)
         RsDetails12("Pand").value = ProjectDes_ID
         RsDetails12("Opr").value = IIf(IsNull(RsDev("id").value), Null, RsDev("id").value)
         RsDetails12("AccountCode").value = astrSplit2tems2(0)
         RsDetails12("EsToal").value = val(astrSplit2tems2(1))
         RsDetails12("value").value = val(astrSplit2tems2(2))
         RsDetails12("Des").value = astrSplit2tems2(3)
        
                         
         RsDetails12.update
         Next J
          End If
      ''//////////////
          
                Next i
    
            End With

            Frame11.Visible = False

    End Select

End Sub

Function calbetprice()
Dim discountvalue As Double
Dim netvalue As Double
Dim Projectvalue As Double
Projectvalue = val(TxtProjectCosts.text)
If val(txt_total_discount) <> 0 Then
discountvalue = val(txt_total_discount)
ElseIf val(TxtDiscountPercentage) <> 0 Then
discountvalue = TxtDiscountPercentage * Projectvalue / 100
End If

total_after_discount.text = Projectvalue - discountvalue
End Function

Private Sub Text13_Change()

End Sub

Private Sub txt_total_discount_KeyUp(KeyCode As Integer, Shift As Integer)
TxtDiscountPercentage.text = 0
calbetprice
End Sub

Private Sub TxtDiscountPercentage_KeyUp(KeyCode As Integer, Shift As Integer)
txt_total_discount = 0
calbetprice
End Sub

Private Sub txtid_LostFocus()
    'Dim StrSQL As String
    'Dim RsTemp As New ADODB.Recordset
    ' StrSQL = "select * From  projects where fullcode='" & DCPreFix.text & (txtid.text) & "'"
    '            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '            If RsTemp.RecordCount > 0 Then
    '
    '                Msg = "this project code already exist"
    '                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '
    '                Exit Sub
    '            End If
End Sub

Private Sub Retrive3(current_opr As String)
    Dim RsDev As ADODB.Recordset
 
    StrSQL = "SELECT  * from opr_expenses where opr_fullcode='" & current_opr & "'"
  
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid3
   
            .Rows = .FixedRows + RsDev.RecordCount
   
            For i = .FixedRows To .Rows - 1
            
                .TextMatrix(i, .ColIndex("ExpensesID")) = IIf(IsNull(RsDev("ExpensesID").value), "", RsDev("ExpensesID").value)
            
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("AccountCode").value), "", RsDev("AccountCode").value)
            
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("AccountName").value), "", RsDev("AccountName").value)
   
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
          
                .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
                RsDev.MoveNext
            Next i

            '  If RsDev.RecordCount > 0 Then
            Me.txt_expenses_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
            '   End If
        End With

    End If

End Sub

Private Sub Retrive2(current_opr As String, _
                     currentqty As Double)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
    'On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Enabled = True
          
    'StrSQL = "select * from Transaction_Details where opr_id=" & current_opr & " and term_id=" & current_terms & " and project_id=" & Me.txt_project_id
 
    StrSQL = " SELECT     TOP 100 PERCENT dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEFDetails.ItemId, dbo.TblProcessDEFDetails.UnitID, dbo.TblProcessDEFDetails.Price, "
    StrSQL = StrSQL & " dbo.TblProcessDEFDetails.Cost, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemCase,"
    StrSQL = StrSQL & " dbo.TblItems.ItemNamee"
    StrSQL = StrSQL & " FROM         dbo.TblProcessDEF INNER JOIN"
    StrSQL = StrSQL & " dbo.TblProcessDEFDetails ON dbo.TblProcessDEF.TblProcessDEFID = dbo.TblProcessDEFDetails.TblProcessDEFID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblUnites ON dbo.TblProcessDEFDetails.UnitID = dbo.TblUnites.UnitID INNER JOIN"
    StrSQL = StrSQL & " dbo.TblItems ON dbo.TblProcessDEFDetails.ItemId = dbo.TblItems.ItemID"
    StrSQL = StrSQL & " WHERE     (dbo.TblProcessDEF.TblProcessDEFID = " & val(current_opr) & ")"

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemId")), "", (RsDetails("ItemId").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemId")), "", Trim(RsDetails("ItemId").value))
            FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Cost")), "", (RsDetails("Cost").value)) * currentqty
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("Valu")) = IIf(IsNull(RsDetails("Cost")), 0, (RsDetails("cost"))) * IIf(IsNull(RsDetails("Price")), 0, (RsDetails("Price").value)) * currentqty
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            RsDetails.MoveNext
        Next Num

    End If

End Sub

Private Sub Retrive2old(current_opr As String)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
    'On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Enabled = True
          
    'StrSQL = "select * from Transaction_Details where opr_id=" & current_opr & " and term_id=" & current_terms & " and project_id=" & Me.txt_project_id
 
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where  (payed is null )  and opr_fullcode='" & current_opr & "'"

    'StrSQL = StrSQL + " where opr_id=" & current_opr & " and term_id=" & current_terms & " and project_id=" & Me.txt_project_id

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("Valu")) = IIf(IsNull(RsDetails("Quantity")), 0, (RsDetails("Quantity").value)) * IIf(IsNull(RsDetails("Price")), 0, (RsDetails("Price").value))
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            RsDetails.MoveNext
        Next Num

    End If

End Sub

Private Sub retrive1(Item_ID As String)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
    'On Error GoTo ErrTrap
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid2.Rows = 2
    VSFlexGrid2.Enabled = True
    txt_opr_total.text = 0
          
    StrSQL = "select * from terms_operations where term_fullcode='" & Item_ID & "'" ' & " and project_id=" & Me.txt_project_id
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid2
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("ProjectDes_ID")) = IIf(IsNull(RsDev("ProjectDes_ID").value), "", RsDev("ProjectDes_ID").value)
                .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsDev("fullcode").value), "", RsDev("fullcode").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("id").value), "", RsDev("id").value)
             '''/
             .TextMatrix(i, .ColIndex("expen")) = IIf(IsNull(RsDev("expen").value), "", RsDev("expen").value)
             .TextMatrix(i, .ColIndex("eq")) = IIf(IsNull(RsDev("eq").value), "", RsDev("eq").value)
            .TextMatrix(i, .ColIndex("emps")) = IIf(IsNull(RsDev("emps").value), "", RsDev("emps").value)
             .TextMatrix(i, .ColIndex("matrials")) = IIf(IsNull(RsDev("matrials").value), "", RsDev("matrials").value)
             
            ''/
            
            
                .TextMatrix(i, .ColIndex("item_id")) = IIf(IsNull(RsDev("item_id").value), "", RsDev("item_id").value)
            
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
            
                .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("id").value), "", RsDev("id").value)
           
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev("total").value), "", RsDev("total").value)
                .TextMatrix(i, .ColIndex("period")) = IIf(IsNull(RsDev("period").value), "", RsDev("period").value)
                .TextMatrix(i, .ColIndex("count")) = IIf(IsNull(RsDev("count").value), "", RsDev("count").value)
            
                .TextMatrix(i, .ColIndex("salary")) = IIf(IsNull(RsDev("salary").value), "", RsDev("salary").value)
 
                .TextMatrix(i, .ColIndex("total_items")) = IIf(IsNull(RsDev("total_items").value), "", RsDev("total_items").value)
                .TextMatrix(i, .ColIndex("total_salary")) = IIf(IsNull(RsDev("total_salary").value), "", RsDev("total_salary").value)
                .TextMatrix(i, .ColIndex("total_expenses")) = IIf(IsNull(RsDev("total_expenses").value), "", RsDev("total_expenses").value)

                .TextMatrix(i, .ColIndex("Symbol")) = IIf(IsNull(RsDev("Symbol").value), "", RsDev("Symbol").value)
            
                .TextMatrix(i, .ColIndex("Pre")) = IIf(IsNull(RsDev("Pre").value), "", RsDev("Pre").value)
            
                .TextMatrix(i, .ColIndex("period1")) = IIf(IsNull(RsDev("period1").value), "", RsDev("period1").value)
            
                .TextMatrix(i, .ColIndex("Earlystartweek")) = IIf(IsNull(RsDev("Earlystartweek").value), "", RsDev("Earlystartweek").value)
            
                .TextMatrix(i, .ColIndex("startweek")) = IIf(IsNull(RsDev("startweek").value), "", RsDev("startweek").value)
            
                .TextMatrix(i, .ColIndex("EarlyEndWeek")) = IIf(IsNull(RsDev("EarlyEndWeek").value), "", RsDev("EarlyEndWeek").value)
            
                .TextMatrix(i, .ColIndex("EndWeek")) = IIf(IsNull(RsDev("EndWeek").value), "", RsDev("EndWeek").value)
            
                .TextMatrix(i, .ColIndex("Critical")) = IIf(IsNull(RsDev("Critical").value), "", RsDev("Critical").value)
          
                .TextMatrix(i, .ColIndex("OPRIDD")) = IIf(IsNull(RsDev("OPRIDD").value), "", RsDev("OPRIDD").value)
            
                .TextMatrix(i, .ColIndex("Actperiod")) = IIf(IsNull(RsDev("Actperiod").value), "", RsDev("Actperiod").value)
            
                .TextMatrix(i, .ColIndex("periodView")) = IIf(IsNull(RsDev("periodView").value), "", RsDev("periodView").value)
            
                .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(RsDev("Qty").value), "", RsDev("Qty").value)
            
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
            
                RsDev.MoveNext
            Next i

            Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        
        End With

    End If
          
    ReLineGrid

End Sub
Sub maxx(Optional ByRef Pand As Integer = 0, Optional ByRef OPR As Integer = 0)
     
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
  Set RsDev = New ADODB.Recordset
    If Pand <> 0 Then
   StrSQL = " select max(Pand) as mx from FoxySerial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   Pand = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "FoxySerial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("Pand").value = Pand
RsDev.update
End If
    If OPR <> 0 Then
   StrSQL = " select max(Opr) as mx from FoxySerial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   OPR = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "FoxySerial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("Opr").value = OPR
RsDev.update
End If
End Sub
Function Checked(Optional Pand As Integer = 0, Optional OPR As Integer = 0) As Boolean
     Checked = False
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
      Set RsDev = New ADODB.Recordset
    If Pand <> 0 Then
   StrSQL = " select * from FoxySerial where pand=" & Pand & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
    If OPR <> 0 Then
  StrSQL = " select * from FoxySerial where Opr=" & OPR & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
End Function

Public Sub retrive(Optional Lngid As Long)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    
        Dim RsDevsub As ADODB.Recordset
    Dim StrSQLsub As String
    
    
    Dim i As Integer
    TxtFillData.text = "T"
    'On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        '    XPTxtCurrent.Caption = 0
        '    XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.find "id=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)

    If Not (IsNull(rs("OpenBalanceDate").value)) Then
        Me.Dtp.value = rs("OpenBalanceDate").value
        Me.Dtp1.value = rs("OpenBalanceDate").value
        Me.Dtp2.value = rs("OpenBalanceDate").value
        Me.Dtp3.value = rs("OpenBalanceDate").value
        Me.Dtp4.value = rs("OpenBalanceDate").value
     
        ' Me.Dtp.Enabled = True
    Else
        Me.Dtp.value = Date
        Me.Dtp1.value = Date
        Me.Dtp2.value = Date
        Me.Dtp3.value = Date
        Me.Dtp4.value = Date
                    
        '   Me.Dtp.Enabled = False
    End If

    If Not IsNull(rs("OpenBalanceType").value) Then
        Me.TxtOpenBalance.text = IIf(IsNull(rs("OpenBalance")), "", Trim(rs("OpenBalance")))

        If rs("OpenBalanceType").value = 0 Then
            OptType(0).value = True
            OptType_Click 0
        ElseIf rs("OpenBalanceType").value = 1 Then
            OptType(1).value = True
            OptType_Click 1
        End If
    
    Else
        Me.TxtOpenBalance.text = 0
        Me.OptType(2).value = True
        OptType_Click 2
    End If

  If IsNull(rs("Pstate").value) Then
Ptype(0).value = True
              
  Else
          If rs("Pstate").value = 0 Then
                
                 Ptype(0).value = True
                Else
                 Ptype(1).value = True
                End If
                
  End If
  
  

    If Not IsNull(rs("OpenBalanceType1").value) Then
        Me.TxtOpenBalance1.text = IIf(IsNull(rs("OpenBalance1")), "", Trim(rs("OpenBalance1")))

        If rs("OpenBalanceType1").value = 0 Then
            OptType1(0).value = True
            OptType1_Click 0
        ElseIf rs("OpenBalanceType1").value = 1 Then
            OptType1(1).value = True
            OptType1_Click 1
        End If
    
    Else
        Me.TxtOpenBalance1.text = 0
        Me.OptType1(2).value = True
        OptType1_Click 2
    End If

    If Not IsNull(rs("OpenBalanceType2").value) Then
        Me.TxtOpenBalance2.text = IIf(IsNull(rs("OpenBalance2")), "", Trim(rs("OpenBalance2")))

        If rs("OpenBalanceType2").value = 0 Then
            OptType2(0).value = True
            OptType2_Click 0
        ElseIf rs("OpenBalanceType2").value = 1 Then
            OptType2(1).value = True
            OptType2_Click 1
        End If
    
    Else
        Me.TxtOpenBalance2.text = 0
        Me.OptType2(2).value = True
        OptType2_Click 2
    End If

    If Not IsNull(rs("OpenBalanceType3").value) Then
        Me.TxtOpenBalance3.text = IIf(IsNull(rs("OpenBalance3")), "", Trim(rs("OpenBalance3")))

        If rs("OpenBalanceType3").value = 0 Then
            OptType3(0).value = True
            OptType3_Click 0
        ElseIf rs("OpenBalanceType3").value = 1 Then
            OptType3(1).value = True
            OptType3_Click 1
        End If
    
    Else
        Me.TxtOpenBalance3.text = 0
        Me.OptType3(2).value = True
        OptType3_Click 3
    End If

    If Not IsNull(rs("OpenBalanceType4").value) Then
        Me.TxtOpenBalance4.text = IIf(IsNull(rs("OpenBalance4")), "", Trim(rs("OpenBalance4")))

        If rs("OpenBalanceType4").value = 0 Then
            OptType4(0).value = True
            OptType4_Click 0
        ElseIf rs("OpenBalanceType4").value = 1 Then
            OptType4(1).value = True
            OptType4_Click 1
        End If
    
    Else
        Me.TxtOpenBalance4.text = 0
        Me.OptType4(2).value = True
        OptType4_Click 4
    End If

    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Enabled = True
    XPTxtSum.text = 0
          
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.Rows = 2
    Fg_Journal.Enabled = True
    txt_total_sum.text = 0
          
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    VSFlexGrid1.Enabled = True
         
            GridSub.Clear flexClearScrollable, flexClearEverything
    GridSub.Rows = 2
    
Me.DcEmp1.BoundText = IIf(IsNull(rs("EmpId1")), "", rs("EmpId1"))
Me.DcEmp.BoundText = IIf(IsNull(rs("EmpId")), "", rs("EmpId"))

    txt_project_id = IIf(IsNull(rs("id").value), 0, val(rs("id").value))

    DcAccount1.BoundText = IIf(IsNull(rs("End_user_Account").value), "", rs("End_user_Account").value)
    DcCurrency.BoundText = IIf(IsNull(rs("CurrencyID").value), 1, rs("CurrencyID").value)

    DTStartDate.value = IIf(IsNull(rs("StartDate").value), Date, rs("StartDate").value)
        DTEnddate.value = IIf(IsNull(rs("Enddate").value), Date, rs("Enddate").value)
    

    Me.DcAccount2.BoundText = IIf(IsNull(rs("End_user_id").value), "", rs("End_user_id").value)

    DcAccount3.BoundText = IIf(IsNull(rs("sub_contractor_Account").value), "", rs("sub_contractor_Account").value)
    DcAccount4.BoundText = IIf(IsNull(rs("sub_contractor_id").value), "", rs("sub_contractor_id").value)

    DCPreFix.text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)

    txtid.text = IIf(IsNull(rs("code").value), "", rs("code").value)

    Me.TXTprojectname.text = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
    Me.TXTprojectnamee.text = IIf(IsNull(rs("Project_namee").value), "", rs("Project_namee").value)

    Me.TxtProjectCosts.text = IIf(IsNull(rs("project_cost").value), 0, rs("project_cost").value)

    Me.txt_total_discount.text = IIf(IsNull(rs("general_discount").value), 0, rs("general_discount").value)
    Me.TxtDiscountPercentage.text = IIf(IsNull(rs("DiscountPercentage").value), 0, rs("DiscountPercentage").value)
    
    Me.total_after_discount.text = IIf(IsNull(rs("cost_after_discount").value), 0, rs("cost_after_discount").value)

    Me.Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.DcbDept.BoundText = IIf(IsNull(rs("Dept_ID").value), "", rs("Dept_ID").value)
    Me.DataCombo1.text = IIf(IsNull(rs("Project_status").value), "", rs("Project_status").value)

    Me.DataCombo5.BoundText = IIf(IsNull(rs("Contract_type").value), "", rs("Contract_type").value)
    Me.txt_total_sum.text = IIf(IsNull(rs("total").value), "", rs("total").value)

    Me.txt_sub_discount.text = IIf(IsNull(rs("sub_discount_total").value), "", rs("sub_discount_total").value)

    Me.txt_sub_net.text = IIf(IsNull(rs("net").value), "", rs("net").value)

    'Me.EXPANSES.text = IIf(IsNull(rs("expanses_account").value), "", rs("expanses_account").value)
    'Me.REVENUE.text = IIf(IsNull(rs("REVENUE_account").value), "", rs("REVENUE_account").value)

    'Me.Material.text = IIf(IsNull(rs("Material_account").value), "", rs("Material_account").value)
    'Me.legal.text = IIf(IsNull(rs("legal").value), "", rs("legal").value)

    'Me.salary.text = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)

    'Me.txtProject_account.text = IIf(IsNull(rs("Project_account").value), "", rs("Project_account").value)

    Me.XPTxtSum.text = IIf(IsNull(rs("items_total").value), 0, rs("items_total").value)

    '»‰Êœ «·„‘—Ê⁄
    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "SELECT    projects_des.fullcode, [index], des, qty, cost, dbo.projects_des.total, discount, net, project_id,sub_contractor_id,CusName , oprid "
        StrSQL = StrSQL + " From dbo.projects_des   LEFT OUTER JOIN  dbo.TblCustemers ON dbo.projects_des.sub_contractor_id = dbo.TblCustemers.CusID "
        StrSQL = StrSQL + " Where (project_id =" & Me.txt_project_id.text & ")"
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   

        If Not (RsDev.BOF Or rs.EOF) Then
            RsDev.MoveFirst
    
            With Me.Fg_Journal
                .Rows = .FixedRows + RsDev.RecordCount

                For i = .FixedRows To .Rows - 1
    
                    .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsDev("fullcode").value), "", RsDev("fullcode").value)
                     .TextMatrix(i, .ColIndex("oprid")) = IIf(IsNull(RsDev("oprid").value), "", RsDev("oprid").value)
                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("des").value), "", RsDev("des").value)
            
                    .TextMatrix(i, .ColIndex("qty")) = IIf(IsNull(RsDev("qty").value), "", RsDev("qty").value)
            
                    .TextMatrix(i, .ColIndex("cost")) = IIf(IsNull(RsDev("cost").value), "", RsDev("cost").value)
           
                    .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev("total").value), "", RsDev("total").value)
        
                    .TextMatrix(i, .ColIndex("discount")) = IIf(IsNull(RsDev("discount").value), "", RsDev("discount").value)
            
                    .TextMatrix(i, .ColIndex("net")) = IIf(IsNull(RsDev("net").value), "", RsDev("net").value)
                    .TextMatrix(i, .ColIndex("net")) = IIf(IsNull(RsDev("net").value), "", RsDev("net").value)
            
                    .TextMatrix(i, .ColIndex("sub_contractor_id")) = IIf(IsNull(RsDev("sub_contractor_id").value), "", RsDev("sub_contractor_id").value)
            
                    .TextMatrix(i, .ColIndex("by")) = IIf(IsNull(RsDev("CusName").value), "", RsDev("CusName").value)
 
                    RsDev.MoveNext
                Next i

                Me.txt_total_sum.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
                Me.txt_sub_discount.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("discount"), .Rows - 1, .ColIndex("discount"))
                Me.txt_sub_net.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("net"), .Rows - 1, .ColIndex("net"))
            End With

        End If
    
    
    
    
    
    
    'œ›⁄«  «·„‘—Ê⁄
        '»‰Êœ «·„‘—Ê⁄
    '-----------------------------------------------------------------------------
 
        StrSQL = "SELECT    * from Projectssub "
 
        StrSQL = StrSQL + " Where (projectid =" & Me.txt_project_id.text & ")"
        Set RsDevsub = New ADODB.Recordset
        RsDevsub.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDevsub.BOF Or RsDevsub.EOF) Then
            RsDevsub.MoveFirst
    
            With Me.GridSub
                .Rows = .FixedRows + RsDevsub.RecordCount

                For i = .FixedRows To .Rows - 1
    
                    .TextMatrix(i, .ColIndex("id")) = i
    
                    .TextMatrix(i, .ColIndex("subdate")) = IIf(Not IsDate(RsDevsub("subdate").value), "", RsDevsub("subdate").value)
            
                   ' .TextMatrix(i, .ColIndex("DesTerm")) = IIf(IsNull(RsDevsub("DesTerm").value), "", RsDevsub("DesTerm").value)
            .TextMatrix(i, .ColIndex("rate")) = IIf(IsNull(RsDevsub("rate").value), "", RsDevsub("rate").value)
                    .TextMatrix(i, .ColIndex("ValueTerm")) = IIf(IsNull(RsDevsub("ValueTerm").value), "", RsDevsub("ValueTerm").value)
           
                    .TextMatrix(i, .ColIndex("SubValue")) = IIf(IsNull(RsDevsub("SubValue").value), "", RsDevsub("SubValue").value)
        
                    .TextMatrix(i, .ColIndex("REmarks")) = IIf(IsNull(RsDevsub("REmarks").value), "", RsDevsub("REmarks").value)
            
            
                    RsDevsub.MoveNext
                Next i

              
            End With

        End If
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        '„Ê«œ «·„‘—Ê⁄
 
        'StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & _
        '"ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
        'StrSQL = StrSQL + " where Project_id=" & Val(txt_project_id.text)

        'Set RsDetails = New ADODB.Recordset
        'RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        'If Not (RsDetails.EOF Or RsDetails.BOF) Then
        '    FG.Rows = RsDetails.RecordCount + 1
        '    For Num = 1 To RsDetails.RecordCount
        '        FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
        '        FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        '        FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
        '        FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
        '        FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
        '        FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
        '        FG.TextMatrix(Num, FG.ColIndex("Valu")) = IIf(IsNull(RsDetails("Quantity")), 0, (RsDetails("Quantity").value)) * IIf(IsNull(RsDetails("Price")), 0, (RsDetails("Price").value))
        '        FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
        '        FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        '        RsDetails.MoveNext
        '    Next Num
        'End If
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
        'EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
        'StrSQL = "SELECT  * FROM  TblEmployee Where (project_id =" & Me.txt_project_id.text & ")"
        '    Set RsDev = New ADODB.Recordset
        '    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        '    If Not (RsDev.BOF Or Rs.EOF) Then
        '      RsDev.MoveFirst
        '
        '    With Me.VSFlexGrid1
        '    .Rows = .FixedRows + RsDev.RecordCount
        '    For I = .FixedRows To .Rows - 1
        '        .TextMatrix(I, .ColIndex("id")) = IIf(IsNull(RsDev("Emp_ID").value), _
        '            "", RsDev("Emp_ID").value)
        '
        '        .TextMatrix(I, .ColIndex("code")) = IIf(IsNull(RsDev("Emp_Code").value), _
        '            "", RsDev("Emp_Code").value)
        '
        '                .TextMatrix(I, .ColIndex("name")) = IIf(IsNull(RsDev("Emp_Name").value), _
        '            "", RsDev("Emp_Name").value)
        '
        '        .TextMatrix(I, .ColIndex("LineNO")) = I
        '        RsDev.MoveNext
        '    Next I
        '    End With
    
        '    End If
        'EEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE
    End If
   
    '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), _
    '  .Rows - 1, .ColIndex("CreditValue"))
    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), _
    '  .Rows - 1, .ColIndex("DebitValue"))

    '-----------------------------------------------------------------------------
    'XPTxtCurrent.Caption = Rs.AbsolutePosition
    'XPTxtCount.Caption = Rs.RecordCount
    ReLineGrid
    Exit Sub
ErrTrap:

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
GridSub.Enabled = False
            Fra(7).Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„‘—Ê⁄«  "
            Else
                Me.Caption = "Projects"
            End If
         
            VSFlexGrid3.Enabled = False
            Ele(2).Enabled = False
            FG.Enabled = False
            Frame3.Enabled = False
            VSFlexGrid2.Editable = flexEDNone
        
          '  Fg_Journal.Editable = flexEDNone
            VSFlexGrid3.Editable = flexEDNone
        
            Me.Command1(0).Enabled = True 'ÃœÌœ
            Me.Command1(3).Enabled = True ' ⁄œÌ·
            Me.Command1(1).Enabled = False 'Õ›Ÿ
            Me.Command1(7).Enabled = True 'Õ–›
            Me.Command1(6).Enabled = False ' —«Ã⁄
            Me.Command1(5).Enabled = True '»ÕÀ
            Me.Command1(2).Enabled = True '„—›ﬁ« 
            Command3.Enabled = True ' ﬁ—Ì—
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
 
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Command1(7).Enabled = False
                Me.Command1(3).Enabled = False
            
            End If
        
        Case "N"
        GridSub.Enabled = True
            Fra(7).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«·„‘—Ê⁄«  (ÃœÌœ)"
            Else
                Me.Caption = " Projects(New Record)"
            End If
        
            VSFlexGrid3.Enabled = True
            Ele(2).Enabled = True
            FG.Enabled = True
            Frame3.Enabled = True
            VSFlexGrid2.Editable = flexEDKbdMouse
        
            Fg_Journal.Editable = flexEDKbdMouse
            VSFlexGrid3.Editable = flexEDKbdMouse
        
            Me.Command1(0).Enabled = False 'ÃœÌœ
            Me.Command1(3).Enabled = False ' ⁄œÌ·
            Me.Command1(1).Enabled = True 'Õ›Ÿ
            Me.Command1(7).Enabled = False 'Õ–›
            Me.Command1(6).Enabled = True ' —«Ã⁄
            Me.Command1(5).Enabled = False '»ÕÀ
            Me.Command1(2).Enabled = True '„—›ﬁ« 
            Command3.Enabled = False ' ﬁ—Ì—
         
        Case "E"
        GridSub.Enabled = True
            Fra(7).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«·„‘—Ê⁄« (  ⁄œÌ· )"
            Else
                Me.Caption = "Projects (Edit Current Record)"
            End If

            VSFlexGrid3.Enabled = True
            Ele(2).Enabled = True
            FG.Enabled = True
            Frame3.Enabled = True
            VSFlexGrid2.Editable = flexEDKbdMouse
        
            Fg_Journal.Editable = flexEDKbdMouse
            VSFlexGrid3.Editable = flexEDKbdMouse
        
            Me.Command1(0).Enabled = False 'ÃœÌœ
            Me.Command1(3).Enabled = False ' ⁄œÌ·
            Me.Command1(1).Enabled = True 'Õ›Ÿ
            Me.Command1(7).Enabled = False 'Õ–›
            Me.Command1(6).Enabled = True ' —«Ã⁄
            Me.Command1(5).Enabled = False '»ÕÀ
            Me.Command1(2).Enabled = True '„—›ﬁ« 
            Command3.Enabled = False ' ﬁ—Ì—
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
  
    End Select

    Exit Sub
ErrTrap:

End Sub
 
Private Sub TxtProjectCosts_Change()
calbetprice
End Sub

Private Sub TXTprojectname_GotFocus()
    SwitchKeyboardLang LANG_ARABIC

End Sub

Private Sub TXTprojectnamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid1

        Select Case .ColKey(Col)
 
            Case "name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
             
                StrSQL = "SELECT  * from TblEmployee Where Emp_id=" & val(StrAccountCode)
                Set rs = Nothing
            
                If StrAccountCode <> "" Then
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        .TextMatrix(Row, .ColIndex("code")) = IIf(IsNull(rs("Emp_Code").value), "", rs("Emp_Code").value)
                    End If
                End If
            
                '.TextMatrix(Row, .ColIndex("id")) = get_Expenses_id(StrAccountCode)
        
            Case "code"
                  
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT  * from TblEmployee Where Emp_Code=" & .TextMatrix(Row, Col)
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
          
                    .TextMatrix(Row, .ColIndex("id")) = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
                    
                    .TextMatrix(Row, .ColIndex("name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                Else
                    .TextMatrix(Row, .ColIndex("id")) = ""
                    .TextMatrix(Row, .ColIndex("name")) = ""
                End If

        End Select
   
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
        txt_employee_count = .Rows - 2
        Me.txt_emp_salary.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
   
    End With

    ReLineGrid
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "name"
                Exit Sub
        End Select

    End With

    VSFlexGrid1.ComboList = ""
End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "name"
                StrSQL = "select * from TblEmployee"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Emp_Name", "Emp_ID")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)

    With VSFlexGrid2
  
        On Error Resume Next
        Dim StrAccountCode As String
        Dim Msg As String
        Dim rs As New ADODB.Recordset
        Dim StrSQL As String
        Dim ClsAcc As New ClsAccounts
        Dim LngRow As Long
        Dim code  As String

        With VSFlexGrid2

            Select Case .ColKey(Col)

                Case "name"
                    code = .ComboData
                    .TextMatrix(Row, .ColIndex("OPRIDD")) = code
                    .TextMatrix(Row, .ColIndex("name")) = .ComboItem
                    .TextMatrix(Row, .ColIndex("qty")) = 1
                    REFillOprData (code)

                Case "qty"
                    code = val(.TextMatrix(Row, .ColIndex("OPRIDD")))
                    REFillOprData (code)
            End Select
      
            If Row = .Rows - 1 Then
                .Rows = .Rows + 1
            End If
  
        End With

        ReLineGrid
 
    End With

    If Me.TxtModFlg <> "E" Then Exit Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    'Grid.TextMatrix(Row, Grid.ColIndex("Code"))
    'Grid.TextMatrix(Row, Grid.ColIndex("Name"))
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid2

        If .ColKey(Col) <> "name" Then
       
            .ComboList = ""
        End If

    End With

End Sub

Private Sub VSFlexGrid2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    With Me.VSFlexGrid2

        Select Case .ColKey(Col)

        Case "expensive"
                                  LngRow = Row

 LngCol = Col
Load FrmExchangeOper
FrmExchangeOper.show vbModal

        Case "equep"
                                  LngRow = Row

 LngCol = Col
Load FrmEquepment
FrmEquepment.show vbModal

        Case "employee"
                                  LngRow = Row

 LngCol = Col
Load FrmEmpOper
FrmEmpOper.show vbModal
                 Case "mat"
                                           LngRow = Row

 LngCol = Col

             ' ItemProductionDate Row, Col, , 1
                Load FrmMatrialsOp
                FrmMatrialsOp.show vbModal

                    
                End Select
                End With
End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim LngItemID As Integer
    Dim MyStrList As String

    With Me.VSFlexGrid2

        Select Case .ColKey(Col)
                       Case "expensive"
.ColComboList(.ColIndex("expensive")) = "..."
        
               Case "equep"
.ColComboList(.ColIndex("equep")) = "..."

        Case "employee"
.ColComboList(.ColIndex("employee")) = "..."

Case "mat"
.ColComboList(.ColIndex("mat")) = "..."

            Case "name"
            
If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = " SELECT     ProcessName, TblProcessDEFID"
Else
                          StrSQL = " SELECT     ProcessNamee, TblProcessDEFID"
      
End If

                StrSQL = StrSQL + " from dbo.TblProcessDEF"
                
                StrSQL = StrSQL + " ORDER BY TblProcessDEFID"
                 
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
               If SystemOptions.UserInterface = ArabicInterface Then

                    MyStrList = .BuildComboList(rs, "ProcessName", "TblProcessDEFID")
                Else
                MyStrList = .BuildComboList(rs, "ProcessNameE", "TblProcessDEFID")
                End If
                    '                    Grid.ColComboList = MyStrList
                    VSFlexGrid2.ColComboList(.ColIndex("name")) = "|" & MyStrList
                Else
                    Cancel = True
                End If
            
        End Select

    End With

End Sub

Function get_opr_details(Strbasedon As String, period As Double, period1 As Double, Optional ByRef StartWeek As Double, Optional ByRef lastweek As Double, Optional ByRef EarlyStartWeek As Double, Optional ByRef Earlylastweek As Double)
    On Error Resume Next
    Dim astrSplitItems() As String
    Dim lastend As Double

    If Strbasedon = "" Then
        StartWeek = 0
        lastweek = period
        EarlyStartWeek = 0
        Earlylastweek = period
    Else

        astrSplitItems = Split(Strbasedon, ",")
        lastend = 0

        For i = 0 To 20

            If lastend < getlastend(astrSplitItems(i)) Then
                lastend = getlastend(astrSplitItems(i))
            End If

        Next i
  
        StartWeek = lastend + period1
        EarlyStartWeek = lastend

        lastweek = StartWeek + period

        Earlylastweek = lastweek - period1

    End If

End Function

Function getlastend(str As String) As Double
    Dim i As Integer

    With Me.VSFlexGrid2

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Symbol")) = str Then
                getlastend = val(.TextMatrix(i, .ColIndex("EndWeek")))

            End If

        Next i

    End With

End Function

Private Sub VSFlexGrid2_Click()

    If Not VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode")) = "" Then
      
        current_opr = VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, VSFlexGrid2.ColIndex("fullcode"))
    End If

End Sub
Private Sub VSFlexGrid3_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)

    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid3

        Select Case .ColKey(Col)
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
  
            Case "value"
                Dim sgl As String
  
        End Select

        Me.txt_expenses_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
   
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub


Private Sub VSFlexGrid3_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid3

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel h= True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub VSFlexGrid3_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid3

        Select Case .ColKey(Col)

            Case "AccountName"
                StrSQL = "select * from Expenses_accounts"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub VSFlexGrid4_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid4
  Select Case .ColKey(Col)
Case "FixedAsset"
Dim rs As ADODB.Recordset
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Namee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
.Rows = .Rows + 1
End Select
End With

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error Resume Next
    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

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

    retrive
    Exit Sub
ErrTrap:
End Sub

