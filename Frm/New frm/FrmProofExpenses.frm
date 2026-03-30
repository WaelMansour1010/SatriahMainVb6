VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmProofExpenses 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14235
   Icon            =   "FrmProofExpenses.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   14235
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4995
      Left            =   0
      TabIndex        =   35
      Top             =   2160
      Width           =   14235
      _cx             =   25109
      _cy             =   8811
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
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmProofExpenses.frx":6852
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
         TabIndex        =   36
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
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmProofExpenses.frx":6AFE
      Left            =   15480
      List            =   "FrmProofExpenses.frx":6B0E
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   21
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
      TabIndex        =   15
      Top             =   0
      Width           =   14505
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   16
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
         ButtonImage     =   "FrmProofExpenses.frx":6B27
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   17
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
         ButtonImage     =   "FrmProofExpenses.frx":6EC1
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   18
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
         ButtonImage     =   "FrmProofExpenses.frx":725B
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   19
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
         ButtonImage     =   "FrmProofExpenses.frx":75F5
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "≈À»«  «·„’—Ê›« "
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
         TabIndex        =   20
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmProofExpenses.frx":798F
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6255
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   14235
      Begin VB.CheckBox ChkAll 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   255
         Left            =   12120
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   840
         Width           =   1935
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   14175
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   8760
            TabIndex        =   12
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   104726529
            CurrentDate     =   38784
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   585
            Index           =   3
            Left            =   120
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   8295
            _cx             =   14631
            _cy             =   1032
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
            Caption         =   " Õœœ «·› —…"
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   6
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   1
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
            Begin VB.ComboBox CmbMonth 
               Height          =   315
               Left            =   75
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   180
               Width           =   2445
            End
            Begin VB.ComboBox CboYear 
               Height          =   315
               Left            =   3915
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   165
               Width           =   2445
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   195
               Index           =   1
               Left            =   2625
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   270
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
               Height          =   240
               Index           =   0
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   180
               Width           =   1020
            End
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   12
            Left            =   7440
            TabIndex        =   45
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   285
            Index           =   4
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   10410
            TabIndex        =   13
            Top             =   255
            Width           =   885
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   585
         Index           =   0
         Left            =   120
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   840
         Width           =   6735
         _cx             =   11880
         _cy             =   1032
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
         Caption         =   " "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   1
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
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmProofExpenses.frx":8D94
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   2895
            _ExtentX        =   5106
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
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   375
            Index           =   0
            Left            =   4800
            TabIndex        =   56
            Top             =   120
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "ﬂ· «·›—Ê⁄"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   375
            Index           =   1
            Left            =   3120
            TabIndex        =   57
            Top             =   120
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "›—⁄ „Õœœ"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   23
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
      TabIndex        =   24
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6720
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
         TabIndex        =   27
         Top             =   480
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   26
         Top             =   1080
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   1
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
            ButtonImage     =   "FrmProofExpenses.frx":8DA9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8880
            TabIndex        =   3
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
            ButtonImage     =   "FrmProofExpenses.frx":F60B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   2
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
            ButtonImage     =   "FrmProofExpenses.frx":F9A5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   4
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
            ButtonImage     =   "FrmProofExpenses.frx":16207
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   5
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
            ButtonImage     =   "FrmProofExpenses.frx":165A1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   1080
            TabIndex        =   6
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
            ButtonImage     =   "FrmProofExpenses.frx":16B3B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   4320
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   240
            Visible         =   0   'False
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
            ButtonImage     =   "FrmProofExpenses.frx":16ED5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   2520
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   240
            Visible         =   0   'False
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
            ButtonImage     =   "FrmProofExpenses.frx":1D737
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   10560
         TabIndex        =   32
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
         Left            =   6960
         TabIndex        =   33
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
         ButtonImage     =   "FrmProofExpenses.frx":1DAD1
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   570
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   -720
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
         Left            =   8640
         TabIndex        =   44
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
         ButtonImage     =   "FrmProofExpenses.frx":24333
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Index           =   3
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«Ã„«·Ì"
         Height          =   210
         Index           =   2
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   13320
         TabIndex        =   34
         Top             =   600
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
            Picture         =   "FrmProofExpenses.frx":2AB95
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProofExpenses.frx":2AF2F
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProofExpenses.frx":2B2C9
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProofExpenses.frx":2B663
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProofExpenses.frx":2B9FD
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProofExpenses.frx":2BD97
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProofExpenses.frx":2C131
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProofExpenses.frx":2C6CB
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   37
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
      ButtonImage     =   "FrmProofExpenses.frx":2CA65
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   40
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
      ButtonImage     =   "FrmProofExpenses.frx":332C7
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   41
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
      ButtonImage     =   "FrmProofExpenses.frx":39B29
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
      TabIndex        =   38
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmProofExpenses"
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

Sub Reline()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.Grid
        For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           Sm = Sm + val(.TextMatrix(i, .ColIndex("Valu")))
           End If
           Next i
  
    End With
   Label2(3).Caption = Sm
End Sub

Private Sub ChkALL_Click()
    Dim i As Integer

    If ChkAll.value = vbChecked Then

        With Me.Grid
        
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("ch")) = True
            Next i

        End With

    Else

        With Me.Grid

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("ch")) = False
            Next i

        End With
         
    End If
    
End Sub

Private Sub Dcbranch_Change()
Dcbranch_Click (0)
End Sub

Private Sub Dcbranch_Click(Area As Integer)
filgrid1
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
       If SystemOptions.UserInterface = ArabicInterface Then
                Grid.ColComboList(Grid.ColIndex("TypeExpens")) = "#1;  Õ”«»|#2; „ÊŸ›"
                ElseIf SystemOptions.UserInterface = EnglishInterface Then
               Grid.ColComboList(Grid.ColIndex("TypeExpens")) = "#1;Account  |#2;Eployee "
               
            End If
            
      YearMonth
      
    conection = "select * from TblProofExpenses order by  ID "
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
    Dcombos.GetBranches Me.Dcbranch
    BtnLast_Click
 
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
Function ChekExpens(Optional ID As Double) As Boolean
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "SELECT     ID, Paye"
sql = sql & " FROM         dbo.TblPripaidExpensesDet"
sql = sql & " where PaymentPayed=1 and id=" & ID & ""
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekExpens = True
Else
ChekExpens = False
End If
End Function
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
    On Error GoTo ErrTrap
    Dim sql As String
    If TxtModFlg = "E" Then
    Remov2All
    'StrSQL = "Delete From TblProofExpensesDet Where ProfExpID='" & val(TxtSerial1.text) & "' "
    'Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    RsSavRec.Fields("RecordM").value = XPDtbTrans.value
    RsSavRec.Fields("BranchId").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("YearID").value = val(Me.CboYear.ListIndex)
    RsSavRec.Fields("MonthID").value = val(Me.CmbMonth.ListIndex)
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    If Rd(1).value = True Then
    RsSavRec.Fields("TypeBranch").value = 1
    Else
    RsSavRec.Fields("TypeBranch").value = 0
    End If
    RsSavRec.update
    ' save grid
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblProofExpensesDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim str2 As String
    With Grid
    
       For i = .FixedRows To .Rows - 1
      If ChekExpens(val(.TextMatrix(i, .ColIndex("IDD")))) = False Then
     If .Cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
                RsDevsub.AddNew
                RsDevsub("ProfExpID").value = val(Me.TxtSerial1.Text)
                RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchID"))) = "", Null, val(.TextMatrix(i, .ColIndex("BranchID"))))
                RsDevsub("TypeExpens").value = IIf((.TextMatrix(i, .ColIndex("TypeExpens"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypeExpens"))))
                RsDevsub("name").value = IIf((.TextMatrix(i, .ColIndex("name"))) = "", Null, .TextMatrix(i, .ColIndex("name")))
                RsDevsub("Account_Code").value = IIf((.TextMatrix(i, .ColIndex("Account_Code"))) = "", Null, (.TextMatrix(i, .ColIndex("Account_Code"))))
                RsDevsub("ExpAccount_Code").value = IIf((.TextMatrix(i, .ColIndex("ExpAccount_Code"))) = "", Null, (.TextMatrix(i, .ColIndex("ExpAccount_Code"))))
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, val(.TextMatrix(i, .ColIndex("EmpID"))))
                RsDevsub("Valuee").value = IIf((.TextMatrix(i, .ColIndex("Valu"))) = "", Null, val(.TextMatrix(i, .ColIndex("Valu"))))
                RsDevsub("IDD").value = IIf((.TextMatrix(i, .ColIndex("IDD"))) = "", Null, val(.TextMatrix(i, .ColIndex("IDD"))))
      RsDevsub.update
      StrSQL = "Update TblPripaidExpensesDet Set  Paye=1, ProfExpID =" & val(TxtSerial1.Text) & " Where ID=" & val(.TextMatrix(i, .ColIndex("IDD"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
      Else
     ' StrSQL = "Update TblPripaidExpensesDet Set  Paye=0  Where ID=" & val(.TextMatrix(i, .ColIndex("IDD"))) & " and ProfExpID =" & val(TxtSerial1.text) & ""
     '      Cn.Execute StrSQL, , adExecuteNoRecords
      End If
      End If
     Next i
    End With
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


' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim i As Integer
    ProgressBar1.Visible = True
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value): ProgressBar1.value = 10
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordM").value), Date, RsSavRec.Fields("RecordM").value): ProgressBar1.value = 20
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchId").value), "", RsSavRec.Fields("BranchId").value): ProgressBar1.value = 30
    CboYear.ListIndex = IIf(IsNull(RsSavRec.Fields("YearID").value), -1, RsSavRec.Fields("YearID").value): ProgressBar1.value = 40
    CmbMonth.ListIndex = IIf(IsNull(RsSavRec.Fields("MonthID").value), -1, RsSavRec.Fields("MonthID").value): ProgressBar1.value = 50
     DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value): ProgressBar1.value = 60
     If Not IsNull(RsSavRec.Fields("TypeBranch").value) Then
     If val(RsSavRec.Fields("TypeBranch").value) = 1 Then
     Rd(1).value = True
     Else
     Rd(0).value = True
     End If
     End If
     LabCurrRec.Caption = RsSavRec.AbsolutePosition: ProgressBar1.value = 70
     LabCountRec.Caption = RsSavRec.RecordCount: ProgressBar1.value = 80
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
    sql = "SELECT     dbo.TblProofExpensesDet.ID, dbo.TblProofExpensesDet.ProfExpID, dbo.TblProofExpensesDet.TypeExpens, dbo.TblProofExpensesDet.name, "
    sql = sql & "                   dbo.TblProofExpensesDet.Valuee, dbo.TblProofExpensesDet.Account_Code, ACCOUNTS_1.Account_Name, ACCOUNTS_1.Account_Serial,"
    sql = sql & "                   ACCOUNTS_1.Account_NameEng, dbo.TblProofExpensesDet.ExpAccount_Code, ACCOUNTS_1.Account_Name AS ExpAccount_Name,"
    sql = sql & "                   ACCOUNTS_1.Account_Serial AS ExpAccount_Serial, ACCOUNTS_1.Account_NameEng AS ExprExpAccount_NameE, dbo.TblProofExpensesDet.EmpID,"
    sql = sql & "                   dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblProofExpensesDet.IDD, dbo.TblProofExpensesDet.BranchID,"
    sql = sql & "                   dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE"
    sql = sql & "  FROM         dbo.TblProofExpensesDet LEFT OUTER JOIN"
    sql = sql & "                    dbo.TblBranchesData ON dbo.TblProofExpensesDet.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblEmployee ON dbo.TblProofExpensesDet.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    sql = sql & "                   dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TblProofExpensesDet.ExpAccount_Code = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
    sql = sql & "                   dbo.ACCOUNTS ACCOUNTS_2 ON dbo.TblProofExpensesDet.Account_Code = ACCOUNTS_2.Account_Code"
    sql = sql & "   Where (dbo.TblProofExpensesDet.ProfExpID =" & val(TxtSerial1.Text) & ")"
   Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
       With Me.Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs1("ID").value), "", Rs1("ID").value)
                   .TextMatrix(i, .ColIndex("IDD")) = IIf(IsNull(Rs1("IDD").value), "", Rs1("IDD").value)
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), "", Rs1("EmpID").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   .TextMatrix(i, .ColIndex("TypeExpens")) = IIf(IsNull(Rs1("TypeExpens").value), 0, Rs1("TypeExpens").value)
                   .TextMatrix(i, .ColIndex("BranchID")) = IIf(IsNull(Rs1("BranchID").value), 0, Rs1("BranchID").value)
                   .TextMatrix(i, .ColIndex("ExpAccount_Code")) = IIf(IsNull(Rs1("ExpAccount_Code").value), "", Rs1("ExpAccount_Code").value)
                   .TextMatrix(i, .ColIndex("Account_Serial1")) = IIf(IsNull(Rs1("ExpAccount_Serial").value), "", Rs1("ExpAccount_Serial").value)
                   
                   .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(Rs1("Account_Code").value), "", Rs1("Account_Code").value)
                   .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(Rs1("Account_Serial").value), "", Rs1("Account_Serial").value)
                   .TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(Rs1("Valuee").value), "", Rs1("Valuee").value)
                .Cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked
                      
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Account_Name1")) = IIf(IsNull(Rs1("ExpAccount_Name").value), "", Rs1("ExpAccount_Name").value)
                   .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(Rs1("Account_Name").value), "", Rs1("Account_Name").value)
                    .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
                   Else
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
                   .TextMatrix(i, .ColIndex("Account_Name1")) = IIf(IsNull(Rs1("ExprExpAccount_NameE").value), "", Rs1("ExprExpAccount_NameE").value)
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs1("Emp_Namee").value), "", Rs1("Emp_Namee").value)
                    .TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(Rs1("Account_NameEng").value), "", Rs1("Account_NameEng").value)
                   
                   End If
                    Rs1.MoveNext
             Next i
            ' .AutoSize 0, .Cols - 1, False
        End With
     Reline
        Exit Sub
 End Sub




Private Sub RemoveGridRow2()
Dim StrSQL As String
Dim i As Integer
Dim k As Integer
If Me.TxtModFlg.Text <> "R" Then

    With Me.Grid
        If Grid.Rows < 2 Then Exit Sub
           k = Grid.Rows - 1
        For i = .FixedRows To Grid.Rows - 1
               
               If ChekExpens(val(.TextMatrix(k, .ColIndex("IDD")))) = False Then
        If .Cell(flexcpChecked, k, .ColIndex("ch")) = flexChecked Then
       StrSQL = "Update TblPripaidExpensesDet Set  Paye=0  Where ID=" & val(.TextMatrix(k, .ColIndex("IDD"))) & " and ProfExpID =" & val(TxtSerial1.Text) & ""
           Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblProofExpensesDet Where IDD=" & val(.TextMatrix(k, .ColIndex("IDD")))
    Cn.Execute StrSQL, , adExecuteNoRecords
        .RemoveItem k
End If
        End If
        k = k - 1
        Next i
    End With

    End If
End Sub

Private Sub RemoveGridRow6()
Dim StrSQL As String
Dim i As Integer
Dim k As Integer
If Me.TxtModFlg.Text <> "R" Then

    With Me.Grid
        If Grid.Rows < 2 Then Exit Sub
           k = Grid.Rows - 1
        For i = .FixedRows To Grid.Rows - 1
               
               If ChekExpens(val(.TextMatrix(k, .ColIndex("IDD")))) = False Then
        If .Cell(flexcpChecked, k, .ColIndex("ch")) = flexChecked Then
       StrSQL = "Update TblPripaidExpensesDet Set  Paye=0  Where ID=" & val(.TextMatrix(k, .ColIndex("IDD"))) & " and ProfExpID =" & val(TxtSerial1.Text) & ""
           Cn.Execute StrSQL, , adExecuteNoRecords
                   StrSQL = "Delete From TblProofExpensesDet Where IDD=" & val(.TextMatrix(k, .ColIndex("IDD")))
    Cn.Execute StrSQL, , adExecuteNoRecords
        .RemoveItem k
End If
        End If
        k = k - 1
        Next i
    End With

    End If
End Sub


Public Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2006 To 3000
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex

End Sub


Sub filgrid1()

   Me.Grid.Clear flexClearScrollable, flexClearEverything
          Grid.Rows = 2
Dim k As Integer
Dim i As Integer
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT     dbo.TblPripaidExpensesDet.ID, dbo.TblPripaidExpensesDet.Name, dbo.TblPripaidExpensesDet.NameE, dbo.TblPripaidExpensesDet.TypeExpens, "
sql = sql & "                      dbo.TblPripaidExpensesDet.EmpID, dbo.TblPripaidExpensesDet.HistoryDate, dbo.TblPripaidExpensesDet.FromDate, dbo.TblPripaidExpensesDet.ToDate,"
sql = sql & "                      dbo.TblPripaidExpensesDet.Valu, dbo.TblPripaidExpensesDet.Remark2, dbo.TblPripaidExpensesDet.Distribution, dbo.TblPripaidExpensesDet.ProofID,"
sql = sql & "                      dbo.TblPripaidExpensesDet.Paye, dbo.TblPripaidExpensesDet.Account_Code, ACCOUNTS_2.Account_Name, ACCOUNTS_2.Account_Serial,"
sql = sql & "                      ACCOUNTS_2.Account_NameEng, dbo.TblPripaidExpensesDet.Account_Code1, ACCOUNTS_1.Account_Name AS ExpAccount_Name,"
sql = sql & "                      ACCOUNTS_1.Account_Serial AS ExpAccount_Serial, ACCOUNTS_1.Account_NameEng AS ExpAccount_NameE, dbo.TblEmployee.Emp_Name,"
sql = sql & "                      dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblPripaidExpensesDet.BranchID, dbo.TblBranchesData.branch_name,"
sql = sql & "                      dbo.TblBranchesData.branch_namee"
sql = sql & " FROM         dbo.TblPripaidExpensesDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblPripaidExpensesDet.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblPripaidExpensesDet.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TblPripaidExpensesDet.Account_Code1 = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
sql = sql & "                      dbo.ACCOUNTS ACCOUNTS_2 ON dbo.TblPripaidExpensesDet.Account_Code = ACCOUNTS_2.Account_Code"
sql = sql & " Where( (dbo.TblPripaidExpensesDet.Paye Is Null)or   (dbo.TblPripaidExpensesDet.Paye = 0)) "
If Me.Dcbranch.Text <> "" And val(Me.Dcbranch.BoundText) <> 0 Then
sql = sql & " and dbo.TblPripaidExpensesDet.BranchID =" & val(Dcbranch.BoundText) & ""
End If
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then

With Grid
.Rows = 2
.Rows = .Rows + Rs8.RecordCount - 1
Rs8.MoveFirst
For k = .FixedRows To Rs8.RecordCount
.TextMatrix(k, .ColIndex("Ser")) = k
.TextMatrix(k, .ColIndex("BranchID")) = IIf(IsNull(Rs8("BranchID").value), "", Rs8("BranchID").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), "", Rs8("branch_name").value)
.TextMatrix(k, .ColIndex("name")) = IIf(IsNull(Rs8("Name").value), "", Rs8("Name").value)
.TextMatrix(k, .ColIndex("Account_Name1")) = IIf(IsNull(Rs8("ExpAccount_Name").value), "", Rs8("ExpAccount_Name").value)
.TextMatrix(k, .ColIndex("Account_Name")) = IIf(IsNull(Rs8("Account_Name").value), "", Rs8("Account_Name").value)
.TextMatrix(k, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Name").value), "", Rs8("Emp_Name").value)
Else
.TextMatrix(k, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
.TextMatrix(k, .ColIndex("name")) = IIf(IsNull(Rs8("NameE").value), "", Rs8("NameE").value)
.TextMatrix(k, .ColIndex("Account_Name1")) = IIf(IsNull(Rs8("ExpAccount_NameE").value), "", Rs8("ExpAccount_NameE").value)
.TextMatrix(k, .ColIndex("Account_Name")) = IIf(IsNull(Rs8("Account_NameEng").value), "", Rs8("Account_NameEng").value)
.TextMatrix(k, .ColIndex("Emp_Name")) = IIf(IsNull(Rs8("Emp_Namee").value), "", Rs8("Emp_Namee").value)
End If
.TextMatrix(k, .ColIndex("TypeExpens")) = IIf(IsNull(Rs8("TypeExpens").value), 0, Rs8("TypeExpens").value)
.TextMatrix(k, .ColIndex("Valu")) = IIf(IsNull(Rs8("Valu").value), 0, Rs8("Valu").value)
.TextMatrix(k, .ColIndex("Account_Code")) = IIf(IsNull(Rs8("Account_Code").value), "", Rs8("Account_Code").value)
.TextMatrix(k, .ColIndex("ExpAccount_Code")) = IIf(IsNull(Rs8("Account_Code1").value), "", Rs8("Account_Code1").value)
.TextMatrix(k, .ColIndex("Account_Serial")) = IIf(IsNull(Rs8("Account_Serial").value), "", Rs8("Account_Serial").value)
.TextMatrix(k, .ColIndex("Account_Serial1")) = IIf(IsNull(Rs8("ExpAccount_Serial").value), "", Rs8("ExpAccount_Serial").value)
.TextMatrix(k, .ColIndex("Fullcode")) = IIf(IsNull(Rs8("Fullcode").value), "", Rs8("Fullcode").value)
.TextMatrix(k, .ColIndex("EmpID")) = IIf(IsNull(Rs8("EmpID").value), 0, Rs8("EmpID").value)
.TextMatrix(k, .ColIndex("IDD")) = IIf(IsNull(Rs8("ID").value), 0, Rs8("ID").value)
Rs8.MoveNext
Next k
.AutoSize 0, .Cols - 1, False
End With
End If
End Sub


Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Reline
End Sub

Private Sub ISButton4_Click()
RemoveGridRow6
End Sub

Private Sub ISButton5_Click()
print_report
End Sub

Private Sub ISButton6_Click()
RemoveGridRow2
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
              
              
              
Dim i As Integer
Dim selectline As Boolean
selectline = False
    With Grid

 

        For i = 1 To .Rows - 1
    
            If .Cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
                           selectline = True
            End If
            
         Next i
         
     End With
     If selectline = False Then
     MsgBox "·„ Ì „  ÕœÌœ œ›⁄« ", vbCritical
     Exit Sub
     End If
              
              
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
    If Rd(1).value = True Then
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
     End If
    ' If val(CboYear.ListIndex) = -1 Then
    '  If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...  Ì—ÃÏ ≈Œ Ì«— «·”‰…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
    '        Dcbranch.SetFocus
    '        Exit Sub
    '        Else
    '        MsgBox "Please Select Year ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        Exit Sub
    '        CboYear.SetFocus
    '     End If
    ' End If
    '      If val(CmbMonth.ListIndex) = -1 Then
    '  If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...  Ì—ÃÏ ≈Œ Ì«— «·‘Â—", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
    '        Dcbranch.SetFocus
    '        Exit Sub
    '        Else
    '        MsgBox "Please Select Month ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        Exit Sub
    '        CmbMonth.SetFocus
    '     End If
    ' End If

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
    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblProofExpenses", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Private Sub Rd_Click(Index As Integer)
If Rd(0).value = True Then
Dcbranch.Enabled = False
Dcbranch.Text = ""
Dcbranch.BoundText = 0
Else
Dcbranch.Enabled = True
End If

filgrid1
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
Sub Remov2All()
Dim i As Integer
Dim StrSQL As String
With Grid
'If Me.TxtModFlg.text <> "R" Then
For i = .FixedRows To .Rows - 1
If ChekExpens(val(.TextMatrix(i, .ColIndex("IDD")))) = False Then
 StrSQL = "Update TblPripaidExpensesDet Set  Paye=0  Where ID=" & val(.TextMatrix(i, .ColIndex("IDD")))
           Cn.Execute StrSQL, , adExecuteNoRecords
           StrSQL = "Delete From TblProofExpensesDet Where IDD=" & val(.TextMatrix(i, .ColIndex("IDD")))
    Cn.Execute StrSQL, , adExecuteNoRecords
    Else
    .Cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked
       
End If
Next i
' cleargriid
' End If
End With
End Sub
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
    Dim i As Integer
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
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
               For i = .FixedRows To .Rows - 1
               If val(.TextMatrix(i, .ColIndex("IDD"))) <> 0 Then
                If ChekExpens(val(.TextMatrix(i, .ColIndex("IDD")))) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… „”œœÂ"
               Else
               MsgBox "Can Not Delete"
               End If
               Exit Sub
               End If
               End If
               Next i
               End With
               Remov2All
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
               '''''''''''''''''''''''''''''''
                 StrSQL = "Delete From TblProofExpensesDet Where ProfExpID='" & val(TxtSerial1.Text) & "'"
                 Cn.Execute StrSQL, , adExecuteNoRecords
    
                

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
   
    TxtModFlg.Text = "N"
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = branch_id
    Dcbranch.SetFocus
     Me.Grid.Clear flexClearScrollable, flexClearEverything
          Grid.Rows = 2
          Rd(1).value = True
           filgrid1
           Label2(3).Caption = 0
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
    Me.Caption = "Proof Expenses  "
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(4).Caption = "Trans ID"
    Me.lbl(2).Caption = "Date"
ELe(3).Caption = "Select Priod"
lbl(0).Caption = "Year"
lbl(1).Caption = "Month"
    Me.lbl(12).Caption = "Branch"
    Rd(0).RightToLeft = False
    Rd(1).RightToLeft = False
    Rd(0).Caption = "All Branch"
    Rd(1).Caption = "Select Branch"
   Label2(2).Caption = "Total"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    ChkAll.RightToLeft = False
    ChkAll.Caption = "Select All"
    '''''''''''''' next
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
        .TextMatrix(0, .ColIndex("name")) = "Name"
        .TextMatrix(0, .ColIndex("TypeExpens")) = "Type Expens"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
        .TextMatrix(0, .ColIndex("Account_Serial")) = "Account Code "
        .TextMatrix(0, .ColIndex("Account_Name")) = "Advance Account"
        .TextMatrix(0, .ColIndex("Account_Name1")) = "Expenses Account"
        .TextMatrix(0, .ColIndex("Valu")) = "Value"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
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
   My_SQL = "TblProofExpenses"
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






