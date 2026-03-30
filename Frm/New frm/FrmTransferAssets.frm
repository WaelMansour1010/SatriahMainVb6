VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTransferAssets 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14235
   Icon            =   "FrmTransferAssets.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   14235
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4395
      Left            =   0
      TabIndex        =   35
      Top             =   2520
      Width           =   14235
      _cx             =   25109
      _cy             =   7752
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmTransferAssets.frx":6852
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
      ItemData        =   "FrmTransferAssets.frx":69D4
      Left            =   15480
      List            =   "FrmTransferAssets.frx":69E4
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
         ButtonImage     =   "FrmTransferAssets.frx":69FD
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
         ButtonImage     =   "FrmTransferAssets.frx":6D97
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
         ButtonImage     =   "FrmTransferAssets.frx":7131
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
         ButtonImage     =   "FrmTransferAssets.frx":74CB
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ‰ﬁ· «·«’Ê·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
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
         Picture         =   "FrmTransferAssets.frx":7865
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
      Begin VB.CheckBox ChkALL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   255
         Left            =   13920
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame d 
         BackColor       =   &H00E2E9E9&
         Height          =   975
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   14175
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2550
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   705
         End
         Begin VB.TextBox TxtNoteIDx 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4440
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   1080
            Width           =   855
         End
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
            Format          =   203161601
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmTransferAssets.frx":8C6A
            Height          =   315
            Left            =   4320
            TabIndex        =   48
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
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
         Begin MSDataListLib.DataCombo DcboEmpName 
            Bindings        =   "FrmTransferAssets.frx":8C7F
            Height          =   315
            Left            =   120
            TabIndex        =   51
            Top             =   240
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
         Begin MSDataListLib.DataCombo DcbFromBranch 
            Bindings        =   "FrmTransferAssets.frx":8C94
            Height          =   315
            Left            =   8760
            TabIndex        =   53
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
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
         Begin MSDataListLib.DataCombo DcbToBranch 
            Bindings        =   "FrmTransferAssets.frx":8CA9
            Height          =   315
            Left            =   4320
            TabIndex        =   54
            Top             =   600
            Width           =   3255
            _ExtentX        =   5741
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
            Caption         =   "„‰ ›—⁄"
            Height          =   285
            Index           =   5
            Left            =   12960
            TabIndex        =   56
            Top             =   600
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï ›—⁄"
            Height          =   285
            Index           =   1
            Left            =   7680
            TabIndex        =   55
            Top             =   600
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁ«∆„  »«·⁄„·Ì…"
            Height          =   285
            Index           =   10
            Left            =   3000
            TabIndex        =   52
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   0
            Left            =   7680
            TabIndex        =   49
            Top             =   240
            Width           =   885
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
         Height          =   705
         Index           =   0
         Left            =   0
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1080
         Width           =   14175
         _cx             =   25003
         _cy             =   1244
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
         Begin VB.TextBox TxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2040
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   61
            Top             =   240
            Width           =   3585
         End
         Begin VB.TextBox TxtAssesetCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11670
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   240
            Width           =   1185
         End
         Begin MSDataListLib.DataCombo DcbAssest 
            Bindings        =   "FrmTransferAssets.frx":8CBE
            Height          =   315
            Left            =   6960
            TabIndex        =   58
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
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
            Height          =   435
            Left            =   120
            TabIndex        =   60
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   767
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
            ButtonImage     =   "FrmTransferAssets.frx":8CD3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   6
            Left            =   5640
            TabIndex        =   62
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«’·"
            Height          =   285
            Index           =   3
            Left            =   12960
            TabIndex        =   59
            Top             =   240
            Width           =   1125
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
      Height          =   2025
      Left            =   0
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6840
      Width           =   14235
      _cx             =   25109
      _cy             =   3572
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
      Begin VB.Frame Frame9 
         Caption         =   "»Ì«‰«  „Õ«”»Ì…"
         Height          =   744
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   600
         Width           =   7896
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   375
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command8 
            Caption         =   "ﬂ‘› Õ”«»"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   195
            Index           =   35
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   27
         Top             =   0
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
         Top             =   1320
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
            ButtonImage     =   "FrmTransferAssets.frx":F535
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
            ButtonImage     =   "FrmTransferAssets.frx":15D97
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
            ButtonImage     =   "FrmTransferAssets.frx":16131
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
            ButtonImage     =   "FrmTransferAssets.frx":1C993
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   5
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   240
            Visible         =   0   'False
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
            ButtonImage     =   "FrmTransferAssets.frx":1CD2D
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
            ButtonImage     =   "FrmTransferAssets.frx":1D2C7
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
            ButtonImage     =   "FrmTransferAssets.frx":1D661
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
            ButtonImage     =   "FrmTransferAssets.frx":23EC3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   10560
         TabIndex        =   32
         Top             =   120
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
         Top             =   120
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
         ButtonImage     =   "FrmTransferAssets.frx":2425D
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
         Top             =   120
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·’› «·Õ«·Ì"
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
         ButtonImage     =   "FrmTransferAssets.frx":2AABF
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   13320
         TabIndex        =   34
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
            Picture         =   "FrmTransferAssets.frx":31321
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferAssets.frx":316BB
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferAssets.frx":31A55
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferAssets.frx":31DEF
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferAssets.frx":32189
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferAssets.frx":32523
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferAssets.frx":328BD
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransferAssets.frx":32E57
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
      ButtonImage     =   "FrmTransferAssets.frx":331F1
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
      ButtonImage     =   "FrmTransferAssets.frx":39A53
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
      ButtonImage     =   "FrmTransferAssets.frx":402B5
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
Attribute VB_Name = "FrmTransferAssets"
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
 Dim ii As Long
 Public LonRow As Double
Public LngCol As Double

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.text, , 200
End Sub

Private Sub DcbAssest_Change()
DcbAssest_Click (0)
End Sub
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords



    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
   
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
        Msg = "  Õ—ﬂ… ‰ﬁ· «’Ê· »—ﬁ„ " & "  »—ﬁ„  " & TxtSerial1
 
 
        Dim FromBranch As Integer
        Dim ToBranch As Integer
      Dim FromBranchAcc As String
        Dim ToBranchAcc As String
 
 
 notes_id = general_noteid

  
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    'C???? C??I?? C?C?C?CE
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim ProjectID As Integer
    Dim group_id As Integer
    Dim Account_code As String
     Dim Account_code2 As String
     
    BranchID = 1
    
    With Grid


line_no = 1
    FromBranch = val(DcbFromBranch.BoundText)
    ToBranch = val(DcbToBranch.BoundText)
       FromBranchAcc = getBranchCurrentAccount(FromBranch)
ToBranchAcc = getBranchCurrentAccount(ToBranch)


        For i = .FixedRows To .rows - 1

    
            If .TextMatrix(i, .ColIndex("FixedID")) <> "" Then      'C?C??? C???E??E IC??
       group_id = val(.TextMatrix(i, .ColIndex("group_id")))
        GetFixedAssetsGroupAccount group_id, , 1, , , , , , Account_code, , Account_code2


            
 If .TextMatrix(i, .ColIndex("PurchasePrice")) > 0 Then
'
           'Account_Code
           '›—⁄ „ÕÊ· «·ÌÂ
           BranchID = ToBranch
           
           StrAccountCode = Account_code
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("PurchasePrice"))), 0, Msg & "  ··«’·    " & .TextMatrix(i, .ColIndex("Fullcode")) & "  " & .TextMatrix(i, .ColIndex("name")) & "„‰ ›—⁄  " & DcbFromBranch.text & " «·Ï ›—⁄ " & DcbToBranch.text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
              StrAccountCode = FromBranchAcc
              If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("PurchasePrice"))), 1, Msg & "  ··«’·    " & .TextMatrix(i, .ColIndex("Fullcode")) & "  " & .TextMatrix(i, .ColIndex("name")) & "„‰ ›—⁄  " & DcbFromBranch.text & " «·Ï ›—⁄ " & DcbToBranch.text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                
                     BranchID = FromBranch
                        StrAccountCode = ToBranchAcc
                        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("PurchasePrice"))), 0, Msg & "  ··«’·    " & .TextMatrix(i, .ColIndex("Fullcode")) & "  " & .TextMatrix(i, .ColIndex("name")) & "„‰ ›—⁄  " & DcbFromBranch.text & " «·Ï ›—⁄ " & DcbToBranch.text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                   
                   
                                 BranchID = FromBranch
                        StrAccountCode = Account_code
                        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("PurchasePrice"))), 1, Msg & "  ··«’·    " & .TextMatrix(i, .ColIndex("Fullcode")) & "  " & .TextMatrix(i, .ColIndex("name")) & "„‰ ›—⁄  " & DcbFromBranch.text & " «·Ï ›—⁄ " & DcbToBranch.text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            
 End If
 '„Ã„⁄ «·«Â·«ﬂ
 
'***********************************************************************************************
 
  If .TextMatrix(i, .ColIndex("AccDepreciation")) > 0 Then
'
           'Account_Code
           '›—⁄ „ÕÊ· «·ÌÂ
           BranchID = ToBranch
           
           StrAccountCode = Account_code2
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("AccDepreciation"))), 1, Msg & "  ··«’·    " & .TextMatrix(i, .ColIndex("Fullcode")) & "  " & .TextMatrix(i, .ColIndex("name")) & "„‰ ›—⁄  " & DcbFromBranch.text & " «·Ï ›—⁄ " & DcbToBranch.text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
              StrAccountCode = FromBranchAcc
              If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("AccDepreciation"))), 0, Msg & "  ··«’·    " & .TextMatrix(i, .ColIndex("Fullcode")) & "  " & .TextMatrix(i, .ColIndex("name")) & "„‰ ›—⁄  " & DcbFromBranch.text & " «·Ï ›—⁄ " & DcbToBranch.text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                
                     BranchID = FromBranch
                   StrAccountCode = Account_code2
                        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("AccDepreciation"))), 0, Msg & "  ··«’·    " & .TextMatrix(i, .ColIndex("Fullcode")) & "  " & .TextMatrix(i, .ColIndex("name")) & "„‰ ›—⁄  " & DcbFromBranch.text & " «·Ï ›—⁄ " & DcbToBranch.text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                   
                   
                                 BranchID = FromBranch
                        StrAccountCode = ToBranchAcc
                        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, val(.TextMatrix(i, .ColIndex("AccDepreciation"))), 1, Msg & "  ··«’·    " & .TextMatrix(i, .ColIndex("Fullcode")) & "  " & .TextMatrix(i, .ColIndex("name")) & "„‰ ›—⁄  " & DcbFromBranch.text & " «·Ï ›—⁄ " & DcbToBranch.text, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            
 End If

'***********************************************************************************************
                
            End If
     
     
     Next i
     
     End With
           
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "Õ—ﬂ… ‰ﬁ· «’Ê· „‰ ›—⁄ «·Ï ›—⁄ »—ﬁ„" & TxtSerial1
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim sql As String
tablename = "TblTransferAssets"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1)
Notevalue = 0

 notytype = 8071
Notevalue = 1 '(total)
 

 BranchID = val(dcBranch.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TXTNoteID.text = NoteID
                                                     TxtNoteSerial.text = NoteSerial
                                     Else
                                                 If TXTNoteID.text = "" Or TxtNoteSerial.text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TXTNoteID.text = NoteID
                                                                TxtNoteSerial.text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TXTNoteID.text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TXTNoteID.text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent
 

     End If

End Function
Private Sub DcbAssest_Click(Area As Integer)
Dim AsseCode1 As String
If val(DcbFromBranch.BoundText) <> 0 Then
If val(DcbAssest.BoundText) <> 0 Then
GetAsseteCode_ID val(DcbAssest.BoundText), AsseCode1, 0
TxtAssesetCode.text = AsseCode1
End If
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— «·›—⁄ «·‰ÕÊ· „‰Â «Ê·«"
Else
MsgBox "Please Select Branch"
End If
DcbFromBranch.SetFocus
Exit Sub
End If
End Sub
Sub GetAsseteCode_ID(Optional ByRef ID As Double = 0, Optional ByRef fullcode As String = "", Optional Typ As Integer = 0)
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If Typ = 0 Then
sql = "select Fullcode  from FixedAssets where id=" & ID & " "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
fullcode = IIf(IsNull(Rs7("Fullcode").value), "", Rs7("Fullcode").value)
Else
fullcode = ""
End If
Else
sql = "select ID  from FixedAssets where Fullcode='" & fullcode & "' "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ID = IIf(IsNull(Rs7("ID").value), 0, Rs7("ID").value)
Else
ID = 0
End If
End If
End Sub

Private Sub DcbAssest_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
If val(Me.DcbFromBranch.BoundText) <> 0 Then
FixedAssetsSearch.branch_no = val(Me.DcbFromBranch.BoundText)
 FixedAssetsSearch.RetrunType = 7
Load FixedAssetsSearch
FixedAssetsSearch.show vbModal
End If
End If
End Sub

Private Sub DcbFromBranch_Change()
DcbFromBranch_Click (0)
End Sub

Private Sub DcbFromBranch_Click(Area As Integer)
loadAssest
End Sub

Private Sub ISButton2_Click()
If DcbFromBranch.text = "" Or val(DcbFromBranch.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄ «·„ÕÊ· „‰Â"
Else
MsgBox "Please Select Branch"
End If
DcbFromBranch.SetFocus
Exit Sub
End If
If DcbToBranch.text = "" Or val(DcbToBranch.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄ «·„ÕÊ· «·ÌÂ"
Else
MsgBox "Please Select Branch"
End If
DcbToBranch.SetFocus
Exit Sub
End If
If DcbAssest.text = "" Or val(DcbAssest.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«—   «·«’·"
Else
MsgBox "Please Select Assets"
End If
DcbAssest.SetFocus
Exit Sub
End If

FillGrid2
End Sub



Private Sub ISButton8_Click()
 FixedAssetsSearch.RetrunType = 10
Load FixedAssetsSearch
FixedAssetsSearch.show vbModal
End Sub

Private Sub TxtAssesetCode_KeyPress(KeyAscii As Integer)
Dim AsseID As Double
If TxtAssesetCode.text <> "" Then
If val(DcbFromBranch.BoundText) <> 0 Then
GetAsseteCode_ID AsseID, TxtAssesetCode.text, 1
DcbAssest.BoundText = AsseID
End If
End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub
Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
End Sub
    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
      
    conection = "select * from TblTransferAssets where 1=1"
    StrSQL = StrSQL & "  AND      BranchID in(" & Current_branchSql & ")"
    
     If SystemOptions.usertype <> UserAdmin Then
     '   conection = conection & " where   BranchID=" & Current_branch
    End If
    conection = conection & "  order by  ID  "
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me
   'load tblUsers -----------------------------------------------
   
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetBranches Me.DcbFromBranch
    Dcombos.GetBranches Me.DcbToBranch
    Dcombos.GetEmployees Me.DcboEmpName
  
    loadAssest
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

Sub loadAssest()
Dim SrSQL As String
   If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " Select  id,Name  From FixedAssets where HaveDepreciation=1   and PurchasePrice>0 and Branch_NO=" & val(Me.DcbFromBranch.BoundText) & ""
        Else
        StrSQL = " Select  id,Namee  From FixedAssets where HaveDepreciation=1   and PurchasePrice>0 and Branch_NO=" & val(Me.DcbFromBranch.BoundText) & ""
        End If
        fill_combo DcbAssest, StrSQL
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
     'On Error GoTo ErrTrap
    Dim sql As String
    If TxtModFlg = "E" Then
    
    StrSQL = "Delete From TblTransferAssetsDet Where TrnsAsseID='" & val(TxtSerial1.text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
              StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords


      End If

    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("BranchID").value = val(Me.dcBranch.BoundText)
    RsSavRec.Fields("FromBranchID").value = val(Me.DcbFromBranch.BoundText)
    RsSavRec.Fields("ToBranchID").value = val(Me.DcbToBranch.BoundText)
    RsSavRec.Fields("EmpID").value = val(Me.DcboEmpName.BoundText)
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.Fields("AssestID").value = val(Me.DcbAssest.BoundText)
    RsSavRec.Fields("Remarks").value = Trim(Me.TxtRemarks.text)
    RsSavRec.update
    ' save grid
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblTransferAssetsDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim str2 As String
    With Grid
    
       For i = .FixedRows To .rows - 1
If val(.TextMatrix(i, .ColIndex("FixedID"))) <> 0 Then
                RsDevsub.AddNew
                RsDevsub("TrnsAsseID").value = val(Me.TxtSerial1.text)
                RsDevsub("FixedID").value = IIf((.TextMatrix(i, .ColIndex("FixedID"))) = "", Null, val(.TextMatrix(i, .ColIndex("FixedID"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remarks"))) = "", Null, .TextMatrix(i, .ColIndex("Remarks")))
                RsDevsub("PurchasePrice").value = IIf((.TextMatrix(i, .ColIndex("PurchasePrice"))) = "", Null, val((.TextMatrix(i, .ColIndex("PurchasePrice")))))
                RsDevsub("AccDepreciation").value = IIf((.TextMatrix(i, .ColIndex("AccDepreciation"))) = "", Null, val((.TextMatrix(i, .ColIndex("AccDepreciation")))))
               RsDevsub.update
     StrSQL = "Update FixedAssets Set  Branch_NO=" & val(Me.DcbToBranch.BoundText) & " Where ID=" & val(.TextMatrix(i, .ColIndex("FixedID"))) & ""
              Cn.Execute StrSQL, , adExecuteNoRecords
        End If
     Next i
    End With
    
    createVoucher
    
      Select Case Me.TxtModFlg.text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
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
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
       
       RsSavRec.Resync adAffectCurrent
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
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value): ProgressBar1.value = 10
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value): ProgressBar1.value = 20
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value): ProgressBar1.value = 30
    Me.DcbFromBranch.BoundText = IIf(IsNull(RsSavRec.Fields("FromBranchID").value), "", RsSavRec.Fields("FromBranchID").value): ProgressBar1.value = 40
    Me.DcbToBranch.BoundText = IIf(IsNull(RsSavRec.Fields("ToBranchID").value), "", RsSavRec.Fields("ToBranchID").value): ProgressBar1.value = 50
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value): ProgressBar1.value = 60
    Me.DcboEmpName.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value): ProgressBar1.value = 70
    Me.DcbAssest.BoundText = IIf(IsNull(RsSavRec.Fields("AssestID").value), "", RsSavRec.Fields("AssestID").value): ProgressBar1.value = 80
    Me.TxtRemarks.text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value): ProgressBar1.value = 90
   Me.TXTNoteID.text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)

  Me.TxtNoteSerial.text = IIf(IsNull(RsSavRec("NoteSerial").value), "", RsSavRec("NoteSerial").value)


     LabCurrRec.Caption = RsSavRec.AbsolutePosition: ProgressBar1.value = 100
     LabCountRec.Caption = RsSavRec.RecordCount: ProgressBar1.value = 10
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
    sql = "SELECT   dbo.TblTransferAssetsDet.ID, dbo.TblTransferAssetsDet.TrnsAsseID, dbo.TblTransferAssetsDet.Remarks, dbo.TblTransferAssetsDet.PurchasePrice, "
    sql = sql & "                   dbo.TblTransferAssetsDet.AccDepreciation , dbo.TblTransferAssetsDet.FixedID, dbo.FixedAssets.name, dbo.FixedAssets.fullcode, dbo.FixedAssets.NameE,dbo.FixedAssets.group_id,dbo.FixedAssets.Branch_NO"
    sql = sql & "   FROM         dbo.TblTransferAssetsDet LEFT OUTER JOIN"
    sql = sql & "                  dbo.FixedAssets ON dbo.TblTransferAssetsDet.FixedID = dbo.FixedAssets.id"
    sql = sql & "   Where (dbo.TblTransferAssetsDet.TrnsAsseID =" & val(TxtSerial1.text) & ")"
   
   Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
       With Me.Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("FixedID")) = IIf(IsNull(Rs1("FixedID").value), 0, Rs1("FixedID").value)
                   .TextMatrix(i, .ColIndex("Branch_NO")) = IIf(IsNull(Rs1("Branch_NO").value), 0, Rs1("Branch_NO").value)
                   .TextMatrix(i, .ColIndex("group_id")) = IIf(IsNull(Rs1("group_id").value), 0, Rs1("group_id").value)
                   
                   .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   .TextMatrix(i, .ColIndex("PurchasePrice")) = IIf(IsNull(Rs1("PurchasePrice").value), 0, Rs1("PurchasePrice").value)
                   .TextMatrix(i, .ColIndex("AccDepreciation")) = IIf(IsNull(Rs1("AccDepreciation").value), 0, Rs1("AccDepreciation").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs1("namee").value), "", Rs1("namee").value)
                   End If
                    Rs1.MoveNext
             Next i
             '.AutoSize 0, .Cols - 1, False
        End With
        Exit Sub
 End Sub




Private Sub RemoveGridRow2()
Dim StrSQL As String
If TxtModFlg.text <> "R" Then
    With Me.Grid
        If Grid.rows < 2 Then Exit Sub
        .RemoveItem .Row
       
    End With
    End If
End Sub

Private Sub ISButton4_Click()
If Me.TxtModFlg.text <> "R" Then
cleargriid
End If
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
                 
                 
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
 
      If dcBranch.text = "" Or val(dcBranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            dcBranch.SetFocus
            Exit Sub
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
             End If
            Exit Sub
            dcBranch.SetFocus
        
     End If
      If Me.DcboEmpName.text = "" Or val(DcboEmpName.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·ﬁ«∆„ »«·⁄„·Ì…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
                  Else
            MsgBox "Please Select Employee ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
             End If
            DcboEmpName.SetFocus
            Exit Sub
     End If
     
    If val(Me.DcbFromBranch.BoundText) = 0 Or DcbFromBranch.text = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...  Ì—ÃÏ ≈Œ Ì«— «·›—⁄ «·„ÕÊ· „‰Â", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
        DcbFromBranch.SetFocus
         Exit Sub
     End If
       If val(Me.DcbToBranch.BoundText) = 0 Or DcbToBranch.text = "" Then
      If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...  Ì—ÃÏ ≈Œ Ì«— «·›—⁄ «·„ÕÊ· «·ÌÂ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
         DcbToBranch.SetFocus
         Exit Sub
     End If
     If val(Me.DcbToBranch.BoundText) = val(Me.DcbFromBranch.BoundText) Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "·«Ì„ﬂ‰ «· ÕÊÌ· ·‰›” «·›—⁄"
     Else
     MsgBox "Can Not Transfer To The Same Branch"
     End If
     Exit Sub
     End If


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
    Select Case Me.TxtModFlg.text
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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblTransferAssets", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "ID=" & RecId, , adSearchForward, 1
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
    FindRec val(TxtSerial1.text)
    Me.TxtModFlg.text = "R"
    FiLLTXT
     BtnLast_Click
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
     If TxtSerial1.text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
       End If
               Else
                          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
               
          
                 StrSQL = "Delete From TblTransferAssetsDet Where TrnsAsseID='" & val(TxtSerial1.text) & "'"
                 Cn.Execute StrSQL, , adExecuteNoRecords
                RsSavRec.Find "ID=" & val(TxtSerial1.text), , adSearchForward, 1
                RsSavRec.delete
               '''''''''''''''''''''''''''''''
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
           Cn.Errors.Clear
    End Select

End Sub
' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
    If TxtModFlg.text = "N" Then
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
       
        ISButton6.Enabled = True
        ISButton4.Enabled = True
        
        
    ElseIf TxtModFlg.text = "R" Then
      '  ISButton6.Enabled = False
        ISButton4.Enabled = False
        
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.text <> "" Then
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
   ElseIf TxtModFlg.text = "E" Then
      '  ISButton6.Enabled = False
        ISButton4.Enabled = False
        
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
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
    If TxtSerial1.text <> "" Then
        TxtModFlg = "E"
        Grid.rows = Grid.rows + 1
        Me.DCboUserName.BoundText = user_id
      '  Me.Dcbranch.BoundText = branch_id
        Frm2.Enabled = True
        Me.dcBranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
 '  Rd(0).value = True
    TxtModFlg.text = "N"
    Me.DCboUserName.BoundText = user_id
    Me.dcBranch.BoundText = Current_branch
    dcBranch.SetFocus
     Me.Grid.Clear flexClearScrollable, flexClearEverything
          Grid.rows = 2
          XPDtbTrans.value = Date
          loadAssest
 
ErrTrap:
End Sub
Sub FillGrid2()
Dim sql As String
Dim k As Integer
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
sql = " SELECT   Branch_NO, group_id,   id, PurchasePrice, PurchaseDate, AccDepreciation"
sql = sql & " From dbo.FixedAssets"
sql = sql & " Where (id =" & val(Me.DcbAssest.BoundText) & ")"
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
With Grid

k = .rows - 1
.rows = .rows + 1

Do While k < (.rows - 1)
.TextMatrix(k, .ColIndex("Ser")) = k

.TextMatrix(k, .ColIndex("FixedID")) = val(DcbAssest.BoundText)
.TextMatrix(k, .ColIndex("name")) = DcbAssest.text
.TextMatrix(k, .ColIndex("Fullcode")) = TxtAssesetCode.text
.TextMatrix(k, .ColIndex("Remarks")) = TxtRemarks.text
.TextMatrix(k, .ColIndex("PurchasePrice")) = IIf(IsNull(Rs9("PurchasePrice").value), 0, Rs9("PurchasePrice").value)
.TextMatrix(k, .ColIndex("AccDepreciation")) = IIf(IsNull(Rs9("AccDepreciation").value), 0, Rs9("AccDepreciation").value)

.TextMatrix(k, .ColIndex("Branch_NO")) = IIf(IsNull(Rs9("Branch_NO").value), 0, Rs9("Branch_NO").value)
.TextMatrix(k, .ColIndex("group_id")) = IIf(IsNull(Rs9("group_id").value), 0, Rs9("group_id").value)

k = k + 1
Loop
.AutoSize 0, .Cols - 1, False
End With
End If
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
 MySQL = " SELECT     dbo.TblTransferAssetsDet.ID, dbo.TblTransferAssetsDet.TrnsAsseID, dbo.TblTransferAssetsDet.Remarks, dbo.TblTransferAssetsDet.PurchasePrice,"
 MySQL = MySQL & "                     dbo.TblTransferAssetsDet.AccDepreciation, dbo.TblTransferAssetsDet.FixedID, dbo.FixedAssets.Name, dbo.FixedAssets.Fullcode, dbo.FixedAssets.namee,"
 MySQL = MySQL & "                     dbo.TblTransferAssets.ID AS MainID, dbo.TblTransferAssets.Remarks AS MainRemarks, dbo.TblTransferAssets.RecordDate, dbo.TblTransferAssets.AssestID,"
 MySQL = MySQL & "                     FixedAssets_1.Name AS MainName, FixedAssets_1.namee AS MainNameE, FixedAssets_1.Fullcode AS MainFullcode, dbo.TblTransferAssets.BranchID,"
 MySQL = MySQL & "                     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblTransferAssets.FromBranchID,"
 MySQL = MySQL & "                     TblBranchesData_1.branch_name AS Frombranch_name, TblBranchesData_1.branch_namee AS Frombranch_nameE, dbo.TblTransferAssets.ToBranchID,"
 MySQL = MySQL & "                     TblBranchesData_2.branch_name AS Tobranch_name, TblBranchesData_2.branch_namee AS Tobranch_nameE, dbo.TblTransferAssets.EmpID,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Name3 , dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality"
 MySQL = MySQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblTransferAssets ON dbo.TblEmployee.Emp_ID = dbo.TblTransferAssets.EmpID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData TblBranchesData_2 ON dbo.TblTransferAssets.ToBranchID = TblBranchesData_2.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData TblBranchesData_1 ON dbo.TblTransferAssets.FromBranchID = TblBranchesData_1.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.FixedAssets FixedAssets_1 ON dbo.TblTransferAssets.AssestID = FixedAssets_1.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblTransferAssets.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblTransferAssetsDet ON dbo.TblTransferAssets.ID = dbo.TblTransferAssetsDet.TrnsAsseID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblTransferAssetsDet.FixedID = dbo.FixedAssets.id"
 MySQL = MySQL & " Where (dbo.TblTransferAssets.id =" & val(TxtSerial1.text) & ")"
  
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransferAssest.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransferAssestE.rpt"
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
       
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
      '  StrReportTitle = "" '& StrAccountName

    Else
 '
 '       xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
      '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
      '  StrReportTitle = ""
    End If
 xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    xReport.ParameterFields(3).AddCurrentValue user_name

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            btnNew_Click
        Else
            Sendkeys "{TAB}"
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
    Me.Caption = "Transfer of Assets   "
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(4).Caption = "Trans ID"
    Me.lbl(2).Caption = "Date"
    Me.lbl(0).Caption = "Branch"
    Me.lbl(5).Caption = "From Branch"
    Me.lbl(1).Caption = "To Branch"
    Me.lbl(10).Caption = "By"
    Me.lbl(3).Caption = "Assets"
    lbl(6).Caption = "Remarks"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    ISButton2.Caption = "Add"
    '''''''''''''' next
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    ISButton6.Caption = "Delet Row"
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
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Code"
        .TextMatrix(0, .ColIndex("name")) = "Assest Name"
        .TextMatrix(0, .ColIndex("PurchasePrice")) = "Value"
        .TextMatrix(0, .ColIndex("AccDepreciation")) = "Acc. Deprec "
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks "
    End With
ErrTrap:
End Sub
Private Sub cleargriid()
Me.Grid.Clear flexClearScrollable, flexClearEverything
Me.Grid.rows = 2
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblTransferAssets"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.text = rs.RecordCount + 1
    Else
        TxtSerial1.text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end



