VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmExpensespaidAdvancedSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ «·„’—Ê›«  «·„œ›Ê⁄… „ﬁœ„«"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13455
   Icon            =   "FrmExpensespaidAdvancedSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   13455
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox Dcbyear 
      Height          =   315
      ItemData        =   "FrmExpensespaidAdvancedSearch.frx":6852
      Left            =   15360
      List            =   "FrmExpensespaidAdvancedSearch.frx":6854
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox DcbMoth 
      Height          =   315
      ItemData        =   "FrmExpensespaidAdvancedSearch.frx":6856
      Left            =   15360
      List            =   "FrmExpensespaidAdvancedSearch.frx":6858
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E2E9E9&
      Height          =   1215
      Left            =   0
      TabIndex        =   33
      Top             =   6720
      Width           =   13455
      Begin VB.OptionButton check3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   12120
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton check4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ÊŸ› „Õœœ"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   10440
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton check5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "›—⁄ „Õœœ"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton check6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈œ«—… „Õœœ…"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   10320
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton check7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„›—œ „Õœœ"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   720
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   39
         Top             =   240
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   40
         Top             =   240
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   6600
         TabIndex        =   41
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   42
         Top             =   720
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton ISButton2 
         Height          =   810
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "»ÕÀ „ ﬁœ„"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1429
         ButtonPositionImage=   3
         Caption         =   "»ÕÀ „ ﬁœ„"
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
         ButtonImage     =   "FrmExpensespaidAdvancedSearch.frx":685A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         LowerToggledContent=   0   'False
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   0
      Width           =   13665
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·»ÕÀ ⁄‰ «·„’—Ê›«  «·„œ›Ê⁄… „ﬁœ„«"
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
         TabIndex        =   26
         Top             =   240
         Width           =   5400
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   12360
         Picture         =   "FrmExpensespaidAdvancedSearch.frx":D0BC
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
         FormatString    =   $"FrmExpensespaidAdvancedSearch.frx":F075
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
      TabIndex        =   19
      Top             =   5400
      Width           =   13455
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   10
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      TabIndex        =   15
      Top             =   6000
      Width           =   13455
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   9480
         TabIndex        =   16
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
         ButtonImage     =   "FrmExpensespaidAdvancedSearch.frx":F28C
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
         TabIndex        =   17
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
         ButtonImage     =   "FrmExpensespaidAdvancedSearch.frx":15AEE
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
         TabIndex        =   18
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
         ButtonImage     =   "FrmExpensespaidAdvancedSearch.frx":1C350
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
      TabIndex        =   10
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”‰œ"
         Height          =   195
         Index           =   14
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   24
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3720
      Width           =   7335
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   3120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   89456643
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   89456643
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·⁄„·Ì…"
         Height          =   195
         Index           =   13
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   315
         Index           =   4
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   315
         Index           =   3
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Height          =   975
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4440
      Width           =   13425
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   1575
         Begin ImpulseButton.ISButton CmdShowMoreOptions 
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ „ ﬁœ„..."
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
            ButtonImage     =   "FrmExpensespaidAdvancedSearch.frx":45F72
            ColorButton     =   14871017
            ColorHoverText  =   12582912
            ButtonToggles   =   1
            DrawFocusRectangle=   0   'False
            RightToLeft     =   -1  'True
            ButtonImageToggled=   "FrmExpensespaidAdvancedSearch.frx":4630C
            ColorToggledHoverText=   192
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   1800
         TabIndex        =   27
         Top             =   120
         Width           =   5055
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ·"
            Height          =   315
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» „⁄Ì‰"
            Height          =   315
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton check2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ÊŸ›Ì‰"
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ—ﬂ…"
            Height          =   285
            Index           =   12
            Left            =   3840
            TabIndex        =   28
            Top             =   240
            Width           =   1005
         End
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Height          =   315
         Left            =   6960
         TabIndex        =   1
         Top             =   240
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboBox 
         Height          =   315
         Left            =   6960
         TabIndex        =   2
         Top             =   600
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lblLL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Left            =   11160
         TabIndex        =   4
         Top             =   240
         Width           =   1965
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„’—Ê› «·„ﬁœ„"
         Height          =   285
         Index           =   7
         Left            =   11160
         TabIndex        =   3
         Top             =   600
         Width           =   2205
      End
   End
End
Attribute VB_Name = "FrmExpensespaidAdvancedSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As FrmExpensespaidAdvancedSearch

Private Sub fg_Click()
FrmExpensespaidAdvanced.FindRec val(Fg.TextMatrix(Fg.Row, 1))
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    AdditemTocCmp
   Set Dcombos = New ClsDataCombos
   Dcombos.GetBranches Me.Dcbranch
   Dcombos.GetExpensespaidAdvanced Me.DcboBox
   Dcombos.GetEmployees Me.DataCombo1
   Dcombos.GetBranches Me.DataCombo2
   Dcombos.GetEmpDepartments Me.DataCombo3
   Dcombos.GetYearlyComponents Me.DataCombo4
   Me.Option1.value = True
   Me.check3.value = True
   HideCulems
    '  Set GrdBack = New ClsBackGroundPic
  '  With Me.Fg
      '  Set .WallPaper = GrdBack.Picture
     '   .AutoSize 0, .Cols - 1, False
   ' End With
   If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
     SetDtpickerDate Me.DtpDateTo
   End Sub
   Private Sub ISButton2_Click()
      If check4.value = True Then
  If DataCombo1.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— «”„  «·„ÊŸ› ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            DataCombo1.SetFocus
             Exit Sub
            DataCombo1.SetFocus
            Else
            MsgBox "Select Employee Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DataCombo1.SetFocus
            Exit Sub
            End If
     End If
  End If
  If check5.value = True Then
  If DataCombo2.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— «”„  «·›—⁄ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            DataCombo2.SetFocus
             Exit Sub
            DataCombo2.SetFocus
            Else
            MsgBox "Select Branch Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DataCombo2.SetFocus
            Exit Sub
            End If
     End If
  End If
 If check6.value = True Then
  If DataCombo3.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— «”„  «·«œ«—… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            DataCombo3.SetFocus
             Exit Sub
            DataCombo3.SetFocus
            Else
            MsgBox "Select Management Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DataCombo3.SetFocus
            Exit Sub
            End If
     End If
  End If
   If check7.value = True Then
   If DataCombo4.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— «·„›—œ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            DataCombo4.SetFocus
             Exit Sub
            DataCombo4.SetFocus
            Else
            MsgBox "Select Single Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DataCombo4.SetFocus
            Exit Sub
            End If
     End If
  End If
      GetDataAdvince
   End Sub
  Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
        GetData
        Case 1
        clear_all Me
          Me.DtpDateFrom.value = ""
          Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lblL(0).Caption = "Search Results"
            End If
         Me.Option1.value = True
         Me.check3.value = True
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
     
    sql = "SELECT   dbo.TbExpensespaidAdvanced.IDEXP, dbo.TbExpensespaidAdvanced.DateM, dbo.TbExpensespaidAdvanced.DateH, dbo.TbExpensespaidAdvanced.BranchID,"
    sql = sql + "     dbo.TbExpensespaidAdvanced.PayWay, dbo.TbExpensespaidAdvanced.Explan, dbo.TbExpensespaidAdvanced.ExpAcount1, dbo.TbExpensespaidAdvanced.ExpAcount,"
    sql = sql + "    dbo.TbExpensespaidAdvanced.ExpSingle, dbo.TbExpensespaidAdvanced.EXPCheck, dbo.TbExpensespaidAdvanced.ExpValue, dbo.TbExpensespaidAdvanced.ExpYear,"
    sql = sql + "    dbo.TbExpensespaidAdvanced.ExpMonth, dbo.TbExpensespaidAdvanced.ExpNumber, dbo.TbExpensespaidAdvanced.ExpEmpCheck, dbo.TbExpensespaidAdvanced.ExpEmpSelect,"
    sql = sql + "     dbo.TbExpensespaidAdvanced.ExpMangemtSelect, dbo.TbExpensespaidAdvanced.ExpSingleSelect, dbo.TbExpensespaidAdvanced.ExpBourchSelect, dbo.TbExpensespaidAdvanced.ExpName,"
    sql = sql + "     dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE, dbo.TbExpensespaidAdvanced.ExpIDD, dbo.TbExpensesprovided.name, dbo.TbExpensesprovided.NameE"
    sql = sql + "      FROM         dbo.TbExpensesprovided RIGHT OUTER JOIN"
    sql = sql + "    dbo.TbExpensespaidAdvanced ON dbo.TbExpensesprovided.ID = dbo.TbExpensespaidAdvanced.ExpIDD LEFT OUTER JOIN"
    sql = sql + "    dbo.TblBranchesData ON dbo.TbExpensespaidAdvanced.BranchID = dbo.TblBranchesData.branch_id"
          
       BolBegine = False
       StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TbExpensespaidAdvanced.IDEXP >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TbExpensespaidAdvanced.IDEXP >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TbExpensespaidAdvanced.IDEXP <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TbExpensespaidAdvanced.IDEXP <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbExpensespaidAdvanced.DateM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TbExpensespaidAdvanced.DateM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.DateM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbExpensespaidAdvanced.DateM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    
    If Me.Dcbranch.text <> "" And (val(Dcbranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.BranchID =" & Me.Dcbranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TbExpensespaidAdvanced.BranchId =" & Me.Dcbranch.BoundText & ""
       End If
     End If
     
     If Me.DcboBox.text <> "" And (val(DcboBox.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND dbo.TbExpensespaidAdvanced.ExpIDD =" & Me.DcboBox.BoundText & ""
        Else
          BolBegine = True
          StrWhere = " Where dbo.TbExpensespaidAdvanced.ExpIDD =" & Me.DcboBox.BoundText & ""
       End If
      End If
     
        If (Me.check1.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.EXPCheck = 1 "
         Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TbExpensespaidAdvanced.EXPCheck = 0 "
        End If
        End If
        
        If (Me.check2.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.EXPCheck = 0 "
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TbExpensespaidAdvanced.EXPCheck = 1 "
        End If
        End If
  '-----------------------------------
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TbExpensespaidAdvanced.IDEXP"
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
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("IDEXP").value), "", rs("IDEXP").value)
                 If Not (IsNull(rs("DateM").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("DateM").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("name").value), "", rs("name").value)
          
               Else
              .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
              .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
               End If
               
               .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(rs("ExpValue").value), "", rs("ExpValue").value)
                             
                Me.DcbMoth.ListIndex = IIf(IsNull(rs.Fields("ExpMonth").value), "-1", rs.Fields("ExpMonth").value)
               .TextMatrix(i, .ColIndex("JobTypeName")) = DcbMoth.text
                                           
                Me.Dcbyear.ListIndex = IIf(IsNull(rs.Fields("ExpYear").value), "-1", rs.Fields("ExpYear").value)
               .TextMatrix(i, .ColIndex("NumEkama")) = Dcbyear.text
               
               .TextMatrix(i, .ColIndex("typevocation")) = IIf(IsNull(rs("ExpNumber").value), "", rs("ExpNumber").value)
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataAdvince()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
     
  sql = "SELECT   dbo.TbExpensespaidAdvanced.DateM, dbo.TbExpensespaidAdvanced.DateH, dbo.TbExpensespaidAdvanced.Explan, dbo.TbExpensespaidAdvanced.ExpIDD,"
  sql = sql + "       dbo.TbExpensespaidAdvanced.PayWay, dbo.TbExpensespaidAdvanced.ExpName, dbo.TbExpensespaidAdvanced.ExpAcount, dbo.TbExpensespaidAdvanced.ExpAcount1,"
  sql = sql + "     dbo.TbExpensespaidAdvanced.ExpSingle, dbo.TbExpensespaidAdvanced.EXPCheck, dbo.TbExpensespaidAdvanced.ExpValue, dbo.TbExpensespaidAdvanced.ExpMonth,"
  sql = sql + "      dbo.TbExpensespaidAdvanced.ExpYear, dbo.TbExpensespaidAdvanced.ExpEmpCheck, dbo.TbExpensespaidAdvanced.ExpNumber, dbo.TbExpensespaidAdvanced.ExpEmpSelect,"
  sql = sql + "     dbo.TbExpensespaidAdvanced.ExpMangemtSelect, dbo.TbExpensespaidAdvanced.ExpBourchSelect, dbo.TbExpensespaidAdvanced.ExpSingleSelect, dbo.TbExpensespaidAdvanced.BranchID,"
  sql = sql + "     dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TbExpensesprovided.ID, dbo.TbExpensesprovided.name,"
  sql = sql + "     dbo.TbExpensesprovided.namee, dbo.TbExpensespaidAdvanced.IDEXP, dbo.TbExpensespaidJoin.IDEXP AS IDEXPHED, dbo.TbExpensespaidJoin.EmpID,"
  sql = sql + "      dbo.TbExpensespaidJoin.BranchID AS BranchIDHED, dbo.TbExpensespaidJoin.MangmentID, dbo.TbExpensespaidJoin.Single, dbo.TbExpensespaidJoin.SingleValue,"
  sql = sql + "      dbo.TbExpensespaidJoin.PayType, dbo.TbExpensespaidJoin.Monthe, dbo.TbExpensespaidJoin.SubYear, dbo.TbExpensespaidJoin.PayValue, dbo.TblEmployee.Emp_ID,"
  sql = sql + "      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Fullcode, TblBranchesData_1.branch_id AS branch_id1, TblBranchesData_1.branch_name AS branch_name1,"
  sql = sql + "      TblBranchesData_1.branch_namee AS branch_namee1, dbo.TblEmpDepartments.DeparmentID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
  sql = sql + "        dbo.mofrdat.mofrad_code , dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee"
  sql = sql + "      FROM         dbo.TblEmployee RIGHT OUTER JOIN"
  sql = sql + "      dbo.TblBranchesData RIGHT OUTER JOIN"
  sql = sql + "        dbo.TbExpensespaidAdvanced LEFT OUTER JOIN"
  sql = sql + "     dbo.mofrdat RIGHT OUTER JOIN"
  sql = sql + "      dbo.TbExpensespaidJoin ON dbo.mofrdat.mofrad_code = dbo.TbExpensespaidJoin.Single LEFT OUTER JOIN"
  sql = sql + "       dbo.TblEmpDepartments ON dbo.TbExpensespaidJoin.MangmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
  sql = sql + "      dbo.TblBranchesData TblBranchesData_1 ON dbo.TbExpensespaidJoin.BranchID = TblBranchesData_1.branch_id ON"
  sql = sql + "       dbo.TbExpensespaidAdvanced.IDEXP = dbo.TbExpensespaidJoin.IDEXP LEFT OUTER JOIN"
  sql = sql + "       dbo.TbExpensesprovided ON dbo.TbExpensespaidAdvanced.ExpIDD = dbo.TbExpensesprovided.ID ON dbo.TblBranchesData.branch_id = dbo.TbExpensespaidAdvanced.BranchID ON"
  sql = sql + "      dbo.TblEmployee.Emp_id = dbo.TbExpensespaidJoin.EmpID"
    
       BolBegine = False
       StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TbExpensespaidAdvanced.IDEXP >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TbExpensespaidAdvanced.IDEXP >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TbExpensespaidAdvanced.IDEXP <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TbExpensespaidAdvanced.IDEXP <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbExpensespaidAdvanced.DateM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TbExpensespaidAdvanced.DateM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.DateM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbExpensespaidAdvanced.DateM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    
    If Me.Dcbranch.text <> "" And (val(Dcbranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.BranchID =" & Me.Dcbranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TbExpensespaidAdvanced.BranchId =" & Me.Dcbranch.BoundText & ""
       End If
     End If
     
     If Me.DcboBox.text <> "" And (val(DcboBox.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND dbo.TbExpensesprovided.ID=" & Me.DcboBox.BoundText & ""
        Else
          BolBegine = True
          StrWhere = " Where dbo.TbExpensesprovided.ID=" & Me.DcboBox.BoundText & ""
       End If
      End If
     
       If (Me.check1.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.EXPCheck = 0 "
         Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TbExpensespaidAdvanced.EXPCheck = 1 "
        End If
        End If
        
        If (Me.check2.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.EXPCheck = 1 "
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TbExpensespaidAdvanced.EXPCheck = 0 "
        End If
        End If
    '-----------------------------------
    ' advinse search
    If (Me.check4.value = True) Then
    If Me.DataCombo1.text <> "" And (val(DataCombo1.BoundText) <> 0) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.ExpEmpSelect = 1 "
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidJoin.EmpID =" & Me.DataCombo1.BoundText & ""
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TbExpensespaidAdvanced.ExpEmpSelect = 1 "
        StrWhere = " Where dbo.TbExpensespaidJoin.EmpID=" & Me.DataCombo1.BoundText & ""
        End If
        End If
        End If
    
      If (Me.check5.value = True) Then
    If Me.DataCombo2.text <> "" And (val(DataCombo2.BoundText) <> 0) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.ExpBourchSelect = 2 "
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidJoin.BranchID =" & Me.DataCombo2.BoundText & ""
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TbExpensespaidAdvanced.ExpBourchSelect = 2 "
        StrWhere = " Where dbo.TbExpensespaidJoin.BranchID=" & Me.DataCombo2.BoundText & ""
        End If
        End If
        End If
        
        
      If (Me.check6.value = True) Then
      If Me.DataCombo3.text <> "" And (val(DataCombo3.BoundText) <> 0) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.ExpMangemtSelect = 3 "
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidJoin.BranchID =" & Me.DataCombo3.BoundText & ""
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TbExpensespaidAdvanced.ExpMangemtSelect = 3 "
        StrWhere = " Where dbo.TbExpensespaidJoin.MangmentID=" & Me.DataCombo3.BoundText & ""
        End If
        End If
        End If
    
      If (Me.check7.value = True) Then
      If Me.DataCombo4.text <> "" And (val(DataCombo4.BoundText) <> 0) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidAdvanced.ExpSingleSelect = 4 "
        StrWhere = StrWhere & " AND  dbo.TbExpensespaidJoin.BranchID =" & Me.DataCombo4.BoundText & ""
        Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TbExpensespaidAdvanced.ExpSingleSelect = 4 "
        StrWhere = " Where dbo.TbExpensespaidJoin.Single=" & Me.DataCombo4.BoundText & ""
        End If
        End If
        End If
    
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TbExpensespaidAdvanced.IDEXP"
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
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("IDEXP").value), "", rs("IDEXP").value)
                 If Not (IsNull(rs("DateM").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("DateM").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("EEmpName")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("EmpBrinch")) = IIf(IsNull(rs("branch_name1").value), "", rs("branch_name1").value)
                .TextMatrix(i, .ColIndex("MangmentEmp")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
                .TextMatrix(i, .ColIndex("EmpSingle")) = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("name").value), "", rs("name").value)
               Else
               .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
               .TextMatrix(i, .ColIndex("EEmpName")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
               .TextMatrix(i, .ColIndex("EmpBrinch")) = IIf(IsNull(rs("branch_namee1").value), "", rs("branch_namee1").value)
               .TextMatrix(i, .ColIndex("MangmentEmp")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
               .TextMatrix(i, .ColIndex("EmpSingle")) = IIf(IsNull(rs("mofrad_namee").value), "", rs("mofrad_namee").value)
               .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
               End If
               
               .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(rs("ExpValue").value), "", rs("ExpValue").value)
                             
                Me.DcbMoth.ListIndex = IIf(IsNull(rs.Fields("ExpMonth").value), "-1", rs.Fields("ExpMonth").value)
               .TextMatrix(i, .ColIndex("JobTypeName")) = DcbMoth.text
                                           
                Me.Dcbyear.ListIndex = IIf(IsNull(rs.Fields("ExpYear").value), "-1", rs.Fields("ExpYear").value)
               .TextMatrix(i, .ColIndex("NumEkama")) = Dcbyear.text
               
               .TextMatrix(i, .ColIndex("typevocation")) = IIf(IsNull(rs("ExpNumber").value), "", rs("ExpNumber").value)
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
     Me.Caption = "Expenses Paid Advanced Search"
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(14).Caption = "Operation ID"
    Me.lbl(5).Caption = "From"
    Me.lbl(6).Caption = "To"
    Me.lbl(13).Caption = "Operation Date"
    Me.lbl(4).Caption = "From"
    Me.lbl(3).Caption = "To"
    Me.lblLL.Caption = "Branch"
    Me.lbl(7).Caption = "Submitted Expense Name"
    Me.lbl(12).Caption = "Movement for"
    Me.check1.Caption = "Select Acount"
    Me.check2.Caption = "Employees"
    Me.CmdShowMoreOptions.Caption = "Advanced"
    Me.Option1.Caption = "All"
    Me.ISButton2.Caption = "Advanced"
    Me.check3.Caption = "All Employees"
    Me.check4.Caption = "Selsct Employee"
    Me.check5.Caption = "Selsct Branch"
    Me.check6.Caption = "Selsct Department"
    Me.check7.Caption = "Selsct Single"
    ''''''''''''''''''''''' next
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "No Transection"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("DepartmentName")) = "DepartmentName"
        .TextMatrix(0, .ColIndex("empname")) = "Emp Name"
       .TextMatrix(0, .ColIndex("GroupName")) = "Location"
         .TextMatrix(0, .ColIndex("JobTypeName")) = "Job Name"
       .TextMatrix(0, .ColIndex("NumEkama")) = "Num Iqama"
       .TextMatrix(0, .ColIndex("typevocation")) = "Type Vocation"
    End With
  '
End Sub
  Private Sub CmdShowMoreOptions_Click()
    If CmdShowMoreOptions.value = True Then
       ShowCulems
       Cmd_Click (1)
       Cmd(0).Enabled = False
       ISButton2.Enabled = True
       Me.Height = Me.Frame7.top + Frame7.Height + 400
       Else
       HideCulems
       Cmd_Click (1)
       Cmd(0).Enabled = True
       ISButton2.Enabled = False
       Me.Height = Me.Frame7.top + 400
    End If
  End Sub
 Private Sub Check3_Click()
 DataCombo1.Enabled = False
  DataCombo1.text = ""
  DataCombo2.Enabled = False
  DataCombo2.text = ""
  DataCombo3.Enabled = False
  DataCombo3.text = ""
  DataCombo4.Enabled = False
  DataCombo4.text = ""
End Sub
Private Sub check4_Click()
If check4.value = vbChecked Then
  DataCombo1.Enabled = False
  Else
  DataCombo1.text = ""
  DataCombo2.text = ""
  DataCombo3.text = ""
  DataCombo4.text = ""
  DataCombo1.Enabled = True
  DataCombo2.Enabled = False
  DataCombo3.Enabled = False
  DataCombo4.Enabled = False
  End If
End Sub
Private Sub check5_Click()
If check5.value = vbChecked Then
  DataCombo2.Enabled = False
  Else
   DataCombo1.text = ""
  DataCombo2.text = ""
  DataCombo3.text = ""
  DataCombo4.text = ""
  DataCombo2.Enabled = True
  DataCombo1.Enabled = False
  DataCombo3.Enabled = False
  DataCombo4.Enabled = False
  End If
End Sub
Private Sub check6_Click()
If check6.value = vbChecked Then
  DataCombo3.Enabled = False
  Else
  DataCombo1.text = ""
  DataCombo2.text = ""
  DataCombo3.text = ""
  DataCombo4.text = ""
  DataCombo3.Enabled = True
  DataCombo1.Enabled = False
  DataCombo2.Enabled = False
  DataCombo4.Enabled = False
  End If
End Sub
Private Sub check7_Click()
If check7.value = vbChecked Then
  DataCombo4.Enabled = False
  Else
  DataCombo1.text = ""
  DataCombo2.text = ""
  DataCombo3.text = ""
  DataCombo4.text = ""
  DataCombo4.Enabled = True
  DataCombo1.Enabled = False
  DataCombo2.Enabled = False
  DataCombo3.Enabled = False
  End If
End Sub
Private Sub AdditemTocCmp()
   Dim i As Integer
  ' full cop month
    For i = 1 To 12
    DcbMoth.AddItem i
    Next i
    
    'full cop year
    For i = 2014 To 2050
    Dcbyear.AddItem i
    Next
    ' full pay way
  End Sub
  Private Sub ShowCulems()
  Fg.ColHidden(4) = False
  Fg.ColHidden(5) = False
  Fg.ColHidden(6) = False
  Fg.ColHidden(7) = False
  End Sub
  Private Sub HideCulems()
   Fg.ColHidden(4) = True
   Fg.ColHidden(5) = True
   Fg.ColHidden(6) = True
   Fg.ColHidden(7) = True
  End Sub
'''''''''''''''''''''''''''' end

