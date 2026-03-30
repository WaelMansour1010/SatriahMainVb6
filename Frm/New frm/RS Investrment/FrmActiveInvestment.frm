VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmActiveInvestment 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14970
   Icon            =   "FrmActiveInvestment.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10785
   ScaleWidth      =   14970
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmActiveInvestment.frx":6852
      Left            =   15480
      List            =   "FrmActiveInvestment.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   76
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
      TabIndex        =   70
      Top             =   0
      Width           =   14985
      Begin VB.TextBox tXTNAME 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   130
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox tXTRootAccount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   129
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   71
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
         ButtonImage     =   "FrmActiveInvestment.frx":687B
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   72
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
         ButtonImage     =   "FrmActiveInvestment.frx":6C15
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   73
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
         ButtonImage     =   "FrmActiveInvestment.frx":6FAF
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   74
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
         ButtonImage     =   "FrmActiveInvestment.frx":7349
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ›⁄Ì· «·„”«Â„…"
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
         TabIndex        =   75
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmActiveInvestment.frx":76E3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   8415
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   720
      Width           =   14955
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·œ›⁄« "
         ForeColor       =   &H000000C0&
         Height          =   3855
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   113
         Top             =   4560
         Width           =   14895
         Begin VB.TextBox TxtSharMetre 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7080
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   3480
            Width           =   1935
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœÊÌ"
            Height          =   252
            Index           =   2
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Œ— ﬁ”ÿ"
            Height          =   252
            Index           =   1
            Left            =   9960
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√Ê· ﬁ”ÿ"
            Height          =   252
            Index           =   0
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox TotalTemp 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   3840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtSharesMeters 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   3840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtSharesCount2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   3480
            Width           =   1815
         End
         Begin VB.TextBox TxtSharesValue2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   3840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtTotalArea 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   10680
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   3480
            Width           =   1935
         End
         Begin VB.TextBox TxtTotalInviseValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   11040
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   3720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   555
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            Top             =   360
            Width           =   4335
         End
         Begin VB.ComboBox DcbPeriodsID 
            Height          =   315
            ItemData        =   "FrmActiveInvestment.frx":8AE8
            Left            =   5640
            List            =   "FrmActiveInvestment.frx":8AEA
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox TxtPeriod 
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
            Left            =   6840
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   360
            Width           =   705
         End
         Begin VB.TextBox TxtPaymentNo 
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
            Left            =   12600
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   360
            Width           =   1065
         End
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   2325
            Left            =   120
            TabIndex        =   114
            Top             =   1080
            Width           =   14685
            _cx             =   25903
            _cy             =   4101
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmActiveInvestment.frx":8AEC
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
         Begin MSComCtl2.DTPicker FristDate 
            Height          =   315
            Left            =   9600
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   93388803
            CurrentDate     =   41640
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   315
            Left            =   5610
            TabIndex        =   38
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
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
            ButtonImage     =   "FrmActiveInvestment.frx":8BB1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSComCtl2.DTPicker TempDate 
            Height          =   270
            Left            =   1560
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   93388803
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ —"
            Height          =   285
            Index           =   30
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   156
            Top             =   3480
            Width           =   435
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Â„ Ì”«ÊÌ"
            Height          =   285
            Index           =   29
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   3480
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   37
            Left            =   12720
            TabIndex        =   150
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«”Â„ ÿ»ﬁ« ··„ —"
            Height          =   285
            Index           =   28
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   3840
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«”Â„"
            Height          =   285
            Index           =   27
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   3480
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·„ —"
            Height          =   285
            Index           =   26
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   3840
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”«Õ… «·«Ã„«·Ì… ··„”«Â„…"
            Height          =   285
            Index           =   25
            Left            =   12840
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   3480
            Width           =   1875
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·‰Â«∆Ì… ··„”«Â„…"
            Height          =   285
            Index           =   24
            Left            =   12360
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   3720
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   21
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·› —Â »Ì‰ «·œ›⁄« "
            Height          =   285
            Index           =   11
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «Ê· œ›⁄Â"
            Height          =   285
            Index           =   9
            Left            =   11160
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·œ›⁄« "
            Height          =   285
            Index           =   8
            Left            =   13800
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   360
            Width           =   930
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   3855
         Left            =   120
         TabIndex        =   69
         Top             =   720
         Width           =   14775
         Begin VB.TextBox TxtInviseNo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox DcbTypePrstg 
            Height          =   315
            ItemData        =   "FrmActiveInvestment.frx":F413
            Left            =   4440
            List            =   "FrmActiveInvestment.frx":F415
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   240
            Width           =   2175
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Height          =   3255
            Left            =   0
            TabIndex        =   99
            Top             =   600
            Width           =   15135
            Begin VB.TextBox TxtRemarks2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   9240
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   170
               Top             =   2040
               Width           =   4575
            End
            Begin VB.TextBox TxtDevelpoValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox TxtLandValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   165
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox TxtAdd 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   1320
               Width           =   1815
            End
            Begin VB.CommandButton Command3 
               Caption         =   "⁄—÷"
               Height          =   315
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   2040
               Width           =   495
            End
            Begin VB.TextBox txtgooglemap 
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
               Left            =   6240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   2040
               Width           =   1455
            End
            Begin VB.ComboBox Dcbownership 
               BackColor       =   &H00FFFFC0&
               Height          =   315
               ItemData        =   "FrmActiveInvestment.frx":F417
               Left            =   120
               List            =   "FrmActiveInvestment.frx":F419
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   240
               Width           =   1695
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÕœÊœ"
               ForeColor       =   &H00C00000&
               Height          =   975
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   2280
               Width           =   14655
               Begin VB.TextBox TxteastWriiten 
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
                  Left            =   3480
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   600
                  Width           =   2145
               End
               Begin VB.TextBox TxtwestWriiten 
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
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   600
                  Width           =   2145
               End
               Begin VB.TextBox TxtPriceSomW 
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
                  Left            =   6960
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   600
                  Width           =   2145
               End
               Begin VB.TextBox TxtPriceHadW 
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
                  Left            =   10320
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   600
                  Width           =   2145
               End
               Begin VB.TextBox txtWestlength 
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
                  Left            =   120
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   240
                  Width           =   2145
               End
               Begin VB.TextBox txtSouthlength 
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
                  Left            =   6960
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   240
                  Width           =   2145
               End
               Begin VB.TextBox txteastlength 
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
                  Left            =   3480
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   240
                  Width           =   2145
               End
               Begin VB.TextBox txtnorthlength 
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
                  Left            =   10320
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   240
                  Width           =   2145
               End
               Begin VB.Label Label23 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘„«·"
                  Height          =   255
                  Left            =   12240
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   600
                  Width           =   855
               End
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘—ﬁ"
                  Height          =   255
                  Left            =   5520
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   600
                  Width           =   855
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ã‰Ê»"
                  Height          =   255
                  Left            =   9240
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   600
                  Width           =   855
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€—»"
                  Height          =   255
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÕœÊœ ﬂ «»Â"
                  Height          =   255
                  Left            =   13200
                  RightToLeft     =   -1  'True
                  TabIndex        =   143
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÕœÊœ «—ﬁ«„"
                  Height          =   255
                  Left            =   13200
                  RightToLeft     =   -1  'True
                  TabIndex        =   142
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€—»"
                  Height          =   255
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ã‰Ê»"
                  Height          =   255
                  Left            =   9240
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘—ﬁ"
                  Height          =   255
                  Left            =   5520
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘„«·"
                  Height          =   255
                  Left            =   12240
                  RightToLeft     =   -1  'True
                  TabIndex        =   138
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.TextBox txtstreetname 
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
               Left            =   120
               MaxLength       =   50
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   20
               Tag             =   "«œŒ· «”„ «·‘«—⁄"
               Top             =   1680
               Width           =   1695
            End
            Begin VB.TextBox Text1 
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
               Left            =   12750
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   600
               Width           =   1065
            End
            Begin VB.TextBox TxtBanckName 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   2400
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TxtPropertyDeed 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   12360
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox TxtIdentityof 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   9240
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox TxtTotalValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   1320
               Width           =   1695
            End
            Begin VB.ComboBox DcbPaymentType 
               Height          =   315
               ItemData        =   "FrmActiveInvestment.frx":F41B
               Left            =   12360
               List            =   "FrmActiveInvestment.frx":F41D
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   1680
               Width           =   1455
            End
            Begin VB.TextBox TxtAddress 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Top             =   2040
               Width           =   4575
            End
            Begin VB.TextBox Text10 
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
               Left            =   12750
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   960
               Width           =   1065
            End
            Begin VB.TextBox TxtMeterValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox TxtArea 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox TxtNo_planned 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TxtDescription 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Top             =   600
               Width           =   4575
            End
            Begin VB.ComboBox DcbType 
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "FrmActiveInvestment.frx":F41F
               Left            =   2880
               List            =   "FrmActiveInvestment.frx":F421
               RightToLeft     =   -1  'True
               TabIndex        =   2
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox TxtSharesValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TxtSharesCount 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   9240
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TxtInviseValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   12360
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   240
               Width           =   1455
            End
            Begin MSDataListLib.DataCombo dcsupplier 
               Height          =   315
               Left            =   9240
               TabIndex        =   10
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   960
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbBank 
               Bindings        =   "FrmActiveInvestment.frx":F423
               Height          =   315
               Left            =   5640
               TabIndex        =   50
               Top             =   2040
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Style           =   2
               BackColor       =   16777088
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
            Begin MSDataListLib.DataCombo DcbLand 
               Height          =   315
               Left            =   9240
               TabIndex        =   5
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   600
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCountryID2 
               Height          =   315
               Left            =   9240
               TabIndex        =   17
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·œÊ·…"
               Top             =   1680
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboGovernmentID 
               Height          =   315
               Left            =   6240
               TabIndex        =   18
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„œÌ‰…"
               Top             =   1680
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCityID 
               Height          =   315
               Left            =   2880
               TabIndex        =   19
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   1680
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcschemeid 
               Height          =   315
               Left            =   12360
               TabIndex        =   49
               Tag             =   " "
               Top             =   1680
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   285
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   169
               Top             =   2040
               Width           =   1275
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «· ÿÊÌ—"
               Height          =   285
               Index           =   32
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   1320
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·«—÷"
               Height          =   285
               Index           =   31
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   960
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… „÷«›…"
               Height          =   285
               Index           =   7
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   159
               Top             =   1320
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„Êﬁ⁄ ÃÊÃ·"
               Height          =   285
               Index           =   33
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   2040
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·„Œÿÿ"
               Height          =   285
               Index           =   15
               Left            =   11040
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   1680
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·‘«—⁄"
               Height          =   285
               Index           =   6
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   1680
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·ÕÌ"
               Height          =   285
               Index           =   5
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   1680
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·„œÌ‰Â"
               Height          =   285
               Index           =   4
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   1680
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·œÊ·…"
               Height          =   285
               Index           =   3
               Left            =   10800
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   1680
               Width           =   1755
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«—÷ „„·Êﬂ…"
               Height          =   285
               Index           =   0
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   600
               Width           =   1275
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "’ﬂ «·„·ﬂÌ…"
               Height          =   285
               Index           =   23
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   1320
               Width           =   1275
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·«—÷"
               Height          =   285
               Index           =   18
               Left            =   1560
               TabIndex        =   120
               Top             =   270
               Width           =   1515
               WordWrap        =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„… «·«Ã„«·Ì…"
               Height          =   285
               Index           =   19
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   1320
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·œ›⁄"
               Height          =   285
               Index           =   11
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   1710
               Width           =   1275
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄‰Ê«‰"
               Height          =   285
               Index           =   17
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   2040
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÂÊÌ…"
               Height          =   285
               Index           =   14
               Left            =   10800
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   1320
               Width           =   1755
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «·„«·ﬂ"
               Height          =   285
               Index           =   1
               Left            =   13560
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   960
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·„ —"
               Height          =   285
               Index           =   13
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   960
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„”«Õ… «·«Ã„«·Ì…"
               Height          =   285
               Index           =   12
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   960
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Œÿÿ"
               Height          =   285
               Index           =   10
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   600
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ê’› «·⁄‰’—"
               Height          =   285
               Index           =   9
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   600
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·⁄‰’—"
               Height          =   285
               Index           =   16
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   270
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·”Â„ «·„»œ∆Ì…"
               Height          =   285
               Index           =   6
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·«”Â„"
               Height          =   285
               Index           =   5
               Left            =   10800
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·„”«Â„…"
               Height          =   285
               Index           =   3
               Left            =   13680
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   240
               Width           =   1275
            End
         End
         Begin VB.TextBox TxtPercenValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   12750
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   705
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Bindings        =   "FrmActiveInvestment.frx":F438
            Height          =   315
            Left            =   8640
            TabIndex        =   42
            Top             =   240
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„”«Â„…"
            Height          =   285
            Index           =   1
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„Ê·… ≈œ«—… «·„”«Â„…"
            Height          =   285
            Index           =   0
            Left            =   6720
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   240
            Width           =   1875
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œÌ— «·„”«Â„…"
            Height          =   285
            Index           =   15
            Left            =   13320
            TabIndex        =   97
            Top             =   240
            Width           =   1365
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         TabIndex        =   66
         Top             =   120
         Width           =   14775
         Begin VB.TextBox TxtInviseOrder 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   12120
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   9240
            TabIndex        =   40
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   93388801
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmActiveInvestment.frx":F44D
            Height          =   315
            Left            =   3480
            TabIndex        =   0
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
            Caption         =   "»‰«¡ ⁄·Ï„”«Â„… —ﬁ„ "
            Height          =   285
            Index           =   22
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   7680
            TabIndex        =   96
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   285
            Index           =   4
            Left            =   13680
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   10890
            TabIndex        =   67
            Top             =   255
            Width           =   885
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
      TabIndex        =   64
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
      TabIndex        =   63
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   78
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
      TabIndex        =   79
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
      Height          =   1665
      Left            =   0
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   9120
      Width           =   14955
      _cx             =   26379
      _cy             =   2937
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
         Height          =   690
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   160
         Top             =   0
         Width           =   4605
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   163
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   161
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
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   164
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   84
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
            TabIndex        =   83
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   81
         Top             =   600
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   57
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
            ButtonImage     =   "FrmActiveInvestment.frx":F462
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8880
            TabIndex        =   59
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
            ButtonImage     =   "FrmActiveInvestment.frx":15CC4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   58
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
            ButtonImage     =   "FrmActiveInvestment.frx":1605E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   60
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
            ButtonImage     =   "FrmActiveInvestment.frx":1C8C0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   61
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
            ButtonImage     =   "FrmActiveInvestment.frx":1CC5A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   62
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
            ButtonImage     =   "FrmActiveInvestment.frx":1D1F4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   3960
            TabIndex        =   94
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
            ButtonImage     =   "FrmActiveInvestment.frx":1D58E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1920
            TabIndex        =   95
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
            ButtonImage     =   "FrmActiveInvestment.frx":23DF0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   9840
         TabIndex        =   87
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
         TabIndex        =   91
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
      Begin ImpulseButton.ISButton ISButton3 
         Height          =   330
         Left            =   3960
         TabIndex        =   157
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
         ButtonImage     =   "FrmActiveInvestment.frx":2418A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   285
         Index           =   20
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   153
         Top             =   240
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   13080
         TabIndex        =   88
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
            Picture         =   "FrmActiveInvestment.frx":2A9EC
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActiveInvestment.frx":2AD86
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActiveInvestment.frx":2B120
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActiveInvestment.frx":2B4BA
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActiveInvestment.frx":2B854
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActiveInvestment.frx":2BBEE
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActiveInvestment.frx":2BF88
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActiveInvestment.frx":2C522
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   89
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
      ButtonImage     =   "FrmActiveInvestment.frx":2C8BC
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   92
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
      ButtonImage     =   "FrmActiveInvestment.frx":3311E
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   93
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
      ButtonImage     =   "FrmActiveInvestment.frx":39980
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
      TabIndex        =   90
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmActiveInvestment"
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
 Dim RevenueAccount As String
 Dim ii As Long
 Public LonRow As Double
Public LngCol As Double

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub DcbLand_Change()

Dim Fullcode As String
If Me.TxtModFlg.Text <> "R" Then
GetInformationBuyLand val(DcbLand.BoundText)
TxtDescription.Text = DcbLand.Text
End If
If val(DcbLand.BoundText) <> 0 Then
GetTblBuyLandRealEstate val(DcbLand.BoundText), Fullcode, 0
If val(Me.Dcbownership.ListIndex) = 1 Then
 TxtDevelpoValue.Text = GetDevelopValue(DcbLand.BoundText)
 End If
Me.Text1.Text = Fullcode
If Me.TxtModFlg.Text <> "R" Then
If CheckLand(val(Me.DcbLand.BoundText)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «” Œœ«„ Â–Â «·«—÷ „‰ ﬁ»·"
Else
MsgBox "Previously it has been using this land"
End If
DcbLand.BoundText = 0
Exit Sub

End If
End If
End If
End Sub
Sub FilrecordLand()
Dim RsDevLand As ADODB.Recordset
Dim StrSQL As String
If Me.TxtModFlg.Text = "E" Then
 StrSQL = "Delete From TblBuyLanReEst Where ActivID =" & val(TxtSerial1.Text) & ""
 Cn.Execute StrSQL, , adExecuteNoRecords
End If
      Set RsDevLand = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblBuyLanReEst Where (1 = -1)"
    RsDevLand.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     Dim StrRecID As String
    StrRecID = new_id("TblBuyLanReEst", "ID", "")
    RsDevLand.AddNew
    RsDevLand.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    RsDevLand.Fields("ActivID").value = val(TxtSerial1.Text)
    RsDevLand.Fields("NewLand").value = 1
    RsDevLand.Fields("Name").value = TxtDescription.Text
    RsDevLand.Fields("RecordDate").value = XPDtbTrans.value
    RsDevLand.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    '''////
    RsDevLand.Fields("CountryID").value = val(Me.DcboCountryID2.BoundText)
    RsDevLand.Fields("CityID").value = val(Me.DcboGovernmentID.BoundText)
    RsDevLand.Fields("HyID").value = val(Me.DcboCityID.BoundText)
   ' RsDevLand.Fields("SchemeID").value = val(Me.dcschemeid.BoundText)
    RsDevLand.Fields("DesLocation").value = (TxtAddress.Text)
    RsDevLand.Fields("Street").value = (txtstreetname.Text)
    RsDevLand.Fields("Northlength").value = val(txtnorthlength.Text)
    RsDevLand.Fields("Southlength").value = val(txtSouthlength.Text)
    RsDevLand.Fields("Eastlength").value = val(txteastlength.Text)
    RsDevLand.Fields("Westlength").value = val((txtWestlength.Text))
    RsDevLand.Fields("NorthlengthStr").value = (TxtPriceHadW.Text)
    RsDevLand.Fields("SouthlengthStr").value = (TxtPriceSomW.Text)
    RsDevLand.Fields("EastlengthStr").value = (TxteastWriiten.Text)
    RsDevLand.Fields("WestlengthStr").value = (TxtwestWriiten.Text)
    RsDevLand.Fields("Googlemap").value = (txtgooglemap.Text)
 ''''//////////////////////
    RsDevLand.Fields("Area").value = val(TxtArea.Text)
    RsDevLand.Fields("FlagActive").value = 1
    'RsDevLand.Fields("No_planned").value = TxtNo_planned.text
    RsDevLand.Fields("MeterPrice").value = val(TxtMeterValue.Text)
    RsDevLand.Fields("OwnerID").value = val(Me.dcsupplier.BoundText)
    RsDevLand.Fields("Total").value = val(Me.TxtTotalValue.Text)
    RsDevLand.Fields("TitledeedNo").value = Me.TxtPropertyDeed.Text
    RsDevLand.Fields("SchemName").value = TxtNo_planned.Text
    RsDevLand.update
    StrSQL = "Update TblActivateInvestment  set LandOwnedID=" & val(StrRecID) & " where id =" & val(TxtSerial1.Text) & ""
    Cn.Execute StrSQL

End Sub
Private Sub DcbLand_Click(Area As Integer)
DcbLand_Change
End Sub
Function CheckLand(Optional ID As Double) As Boolean
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Dim sql As String
sql = "Select id from TblBuyLanReEst where ID =" & ID & " and (FlagActive =1) "
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
CheckLand = True
Else
CheckLand = False
End If
End Function
Private Sub DcbLand_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearchinvestment.inde = 8
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End If
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
    
Private Sub Dcbownership_Change()

If Me.TxtModFlg.Text <> "R" Then
GetInformationBuyLand val(DcbLand.BoundText)
End If
TxtDescription.Text = DcbLand.Text
If val(Me.Dcbownership.ListIndex) = 1 Then
Label1(0).Visible = True
Text1.Visible = True
DcbLand.Visible = True
TxtDevelpoValue.Visible = True
Label1(7).Visible = True
Me.TxtAdd.Visible = True
lbl(32).Visible = True
  Dim Dcombos As New ClsDataCombos
  If Me.TxtModFlg.Text <> "R" Then
    Dcombos.GetBuyLandRealEstate DcbLand, 1, val(TxtSerial1.Text)
    Else
    Dcombos.GetBuyLandRealEstate DcbLand
  End If

Else
Label1(0).Visible = False
Text1.Visible = False
TxtDevelpoValue.Visible = False
DcbLand.Visible = False
DcbLand.BoundText = 0
Label1(7).Visible = False
Me.TxtAdd.Visible = False
TxtAdd.Text = 0
lbl(32).Visible = False
TxtDevelpoValue.Text = 0
End If
End Sub

Private Sub Dcbownership_Click()
Dcbownership_Change
End Sub

Private Sub Dcbranch_Change()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.Text = ""
End If
End Sub

Private Sub DcbType_Change()
If val(Me.DcbType.ListIndex) = 1 Then
lbl(9).Caption = "Ê’› «·⁄ﬁ«—"
Else
lbl(9).Caption = "Ê’› «·«—÷"
End If
End Sub

Private Sub DcbType_Click()
DcbType_Change
End Sub

Private Sub dcsupplier_Change()
dcsupplier_Click (0)
End Sub

Private Sub dcsupplier_Click(Area As Integer)
   If val(dcsupplier.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcsupplier.BoundText, EmpCode
    Me.Text10.Text = EmpCode
  
End Sub
Private Sub LoadDataCombos(Optional BolExceptCountries As Boolean = False, _
                           Optional BolExceptGovern As Boolean = False, _
                           Optional BolExceptCities As Boolean = False)
    Dim Dcombo As New ClsDataCombos
    Dcombo.GetCountriesNames Me.DcboCountryID2
    If BolExceptGovern = False Then
        Dcombo.getCountriesGovernments Me.DcboGovernmentID, val(Me.DcboCountryID2.BoundText)
    End If

    If BolExceptCities = False Then
        Dcombo.GetCountriesGovernCities Me.DcboCityID, val(Me.DcboCountryID2.BoundText), val(Me.DcboGovernmentID.BoundText)
    End If
End Sub
Private Sub DcboCountryID2_Change()
  Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
         'Dcombos.GetCountriesGovernCities Me.DcboCityID, val(Me.DcboCountryID2.BoundText), val(Me.DcboGovernmentID.BoundText)
     Dcombos.getCountriesGovernments Me.DcboGovernmentID, val(Me.DcboCountryID2.BoundText)
End Sub

Private Sub DcboCountryID2_Click(Area As Integer)
DcboCountryID2_Change
End Sub
Private Sub DcboGovernmentID_Change()

    LoadDataCombos False, True, False

End Sub

Private Sub DcboGovernmentID_Click(Area As Integer)
    DcboGovernmentID_Change
End Sub

Private Sub DcboGovernmentID_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
loadcombo
End If
End Sub
Private Sub DcboCountryID2_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF5 Then
loadcombo
End If
End Sub
Sub loadcombo()
 Dim Dcombos As ClsDataCombos
 Dim My_SQL As String
   Set Dcombos = New ClsDataCombos
    Dcombos.get«hay Me.DcboCityID
     Dcombos.getSchemes Me.dcschemeid

End Sub
    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
 
LoadDataCombos
loadcombo
  If SystemOptions.UserInterface = ArabicInterface Then
     With DcbType
       .Clear
       .AddItem "«—«÷Ì"
       .AddItem "⁄ﬁ«—"
    End With
    With Dcbownership
    .Clear
    .AddItem "ÃœÌœ"
    .AddItem "«—÷ „„·Êﬂ…"
    End With
    With DcbTypePrstg
    .Clear
    .AddItem "ﬁÌ„…"
    .AddItem "‰”»…"
    .AddItem "»·«"
    End With
    With DcbPeriodsID
    .Clear
    .AddItem "ÌÊ„"
    .AddItem "‘Â—"
    .AddItem "”‰…"
    End With
    With DcbPaymentType
    .Clear
    .AddItem "ÕÊ«·…"
    .AddItem "‰ﬁœÌ"
    .AddItem "‘Ìﬂ"
    End With
 Else
    With DcbPaymentType
    .Clear
    .AddItem "Transfer"
    .AddItem "Cash"
    .AddItem "Cheque"
    End With
   With DcbTypePrstg
    .Clear
    .AddItem "Value"
    .AddItem "Percentage"
    End With
    
     With DcbPeriodsID
    .Clear
    .AddItem "Day"
    .AddItem "Month"
    .AddItem "Year"
    End With
    With DcbType
      .Clear
      .AddItem "Land"
      .AddItem "Estate"
      '.AddItem "Land owned"
   End With
      With Dcbownership
    .Clear
    .AddItem "New"
    .AddItem "Land Owned"
    End With
End If


    conection = "select * from TblActivateInvestment order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
 
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetBanks Me.DcbBank
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBuyLandRealEstate DcbLand
    Dcombos.GetCustomersSuppliers 2, Me.dcsupplier
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
ErrTrap:
End Sub
Function CheckExp(Optional ID As Double = 0) As Boolean
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
CheckExp = False
sql = "SELECT    InviseNo "
sql = sql & " From dbo.TblActivateInvestment"
sql = sql & " Where   InviseNo=" & ID & " and ExPayed=1"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
CheckExp = True
Else
CheckExp = False
End If
End Function

' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
 
des = "  ›⁄Ì· «·„”«Â„…   »—ﬁ„ " & TxtSerial1 & "  ··„«·ﬂ  " & dcsupplier.Text
 notytype = 9003
 

Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim sql As String
tablename = "TblActivateInvestment"
Filedname = "ID"
NoteSerial1 = TxtSerial1.Text
Notevalue = val(TxtTotalValue.Text)
 

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
      '  Msg = EleHeader.Caption & " ??? " & txtID & " EEC??I" & Date
    If TxtRemarks2.Text = "" Then
Msg = "  ›⁄Ì· «·„”«Â„…   »—ﬁ„ " & TxtSerial1 & "  ··„«·ﬂ  " & dcsupplier.Text
Else
Msg = TxtRemarks2.Text
End If
 
        
 
 notes_id = general_noteid

  
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    'C???? C??I?? C?C?C?CE
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim ProjectID As Integer
    
    BranchID = 1
    

Dim DebitAcc As String
Dim CreditAcc As String
If Dcbownership.ListIndex = 0 Then
DebitAcc = RsSavRec("Account_Code3").value
CreditAcc = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcsupplier.BoundText))
ElseIf Dcbownership.ListIndex = 1 Then
DebitAcc = RsSavRec("Account_Code3").value
 
CreditAcc = GetMyAccountCode("TblBuyLanReEst", "id", val(Me.DcbLand.BoundText))


End If


line_no = 1
 
    BranchID = val(Dcbranch.BoundText)
    
   '  RevenueAccount
   
                 If ModAccounts.AddNewDev(LngDevID, line_no, DebitAcc, val(TxtTotalValue), 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                'CreditAcc
                If ModAccounts.AddNewDev(LngDevID, line_no, CreditAcc, val(TxtTotalValue) - val(TxtAdd), 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
            If val(TxtAdd) > 0 Then
            line_no = line_no + 1
                  If ModAccounts.AddNewDev(LngDevID, line_no, RevenueAccount, val(TxtAdd), 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
            End If
             
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function

Public Sub FiLLRec()
  
  
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                 StrSQL = "Delete From TblActivateInvestmentDet Where ActInvID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                  sql = "update TblBuyLanReEst set FlagActive=null  where ActivID=" & val(TxtSerial1.Text) & ""
                  Cn.Execute sql
                  
                           StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords


              End If
              
            Dim ParentAccount1 As String
            Dim accName As String
             Dim accNameE As String
            ' Dim RootAccount1  As String
            ' Dim RootAccount2  As String
             
                         
     accName = "„”«Â„…  " & DcbType.Text & " " & TxtDescription.Text
                 accNameE = accName
                      
              If Me.TxtModFlg.Text = "N" Then

        'tXTRootAccount Õ”«» «·„”«Â„… «·—∆Ì”Ì
           ParentAccount1 = ModAccounts.AddNewAccount(tXTRootAccount, accName, False, False, accNameE)      'ParentAccount
            RsSavRec("ParentAccount1").value = ParentAccount1 ' Õ”«» «·«—÷
                        RsSavRec("Account_Code3").value = ModAccounts.AddNewAccount(ParentAccount1, accName, True, False, accNameE)
                         RsSavRec("Account_Code4").value = ModAccounts.AddNewAccount(ParentAccount1, accName & "  „’«—Ì›  ÿÊÌ—", True, False, accName & "Development Expenses")
                         
                         RsSavRec("Account_Code5").value = ModAccounts.AddNewAccount(RootAccount1, accName & "  «Ì—«œ« ", True, False, accName & "Revenue")
                         RsSavRec("Account_Code6").value = ModAccounts.AddNewAccount(RootAccount2, accName & "  «—»«Õ ", True, False, accName & "Profit")
                         If RootAccount2 <> RootAccount3 Then '·Ê «·Õ”«» Ê«Õœ ·«  ‰‘∆ Œ” «∆—
                         RsSavRec("Account_Code7").value = ModAccounts.AddNewAccount(RootAccount3, accName & "  Œ”«∆— ", True, False, accName & "Loss")
                         End If
            Else 'edit
    accName = "„”«Â„…  " & DcbType.Text & " " & TxtDescription.Text
                 accNameE = accName
                 If Not IsNull(RsSavRec("ParentAccount1").value) Then
                    ModAccounts.EditAccount RsSavRec("ParentAccount1").value, accName, accNameE, , , , , , , , , , , , , , , , False
                End If
 
                
      If Not IsNull(RsSavRec("Account_Code3").value) Then
                    ModAccounts.EditAccount RsSavRec("Account_Code3").value, accName & "  ", accNameE, , , , , , , , , , , , , , , , True, True
                End If
                
                
                
                
                If Not IsNull(RsSavRec("Account_Code4").value) Then
                    ModAccounts.EditAccount RsSavRec("Account_Code4").value, accName & "  „’«—Ì›  ÿÊÌ—", accNameE & "Dev. Expenses ", , , , , , , , , , , , , , , , True, True
                End If
                
                 
                 
                       If Not IsNull(RsSavRec("Account_Code5").value) Then
                    ModAccounts.EditAccount RsSavRec("Account_Code5").value, accName & "    «Ì—«œ«  ", accNameE & "Revenue", , , , , , , , , , , , , , , , True, True
                End If
                
                
               If Not IsNull(RsSavRec("Account_Code6").value) Then
                    ModAccounts.EditAccount RsSavRec("Account_Code6").value, accName & "    «—»«Õ ", accNameE & "Profit", , , , , , , , , , , , , , , , True, True
                End If
                
                
                          If Not IsNull(RsSavRec("Account_Code7").value) Then
                    ModAccounts.EditAccount RsSavRec("Account_Code7").value, accName & "    Œ”«∆— ", accNameE & "Loss", , , , , , , , , , , , , , , , True, True
                End If
                
                
            
            End If
      ''//////////
      RsSavRec.Fields("Remarks2").value = (TxtRemarks2.Text)
      RsSavRec.Fields("CountryID").value = val(Me.DcboCountryID2.BoundText)
    RsSavRec.Fields("CityID").value = val(Me.DcboGovernmentID.BoundText)
    RsSavRec.Fields("HyID").value = val(Me.DcboCityID.BoundText)
    RsSavRec.Fields("SchemeID").value = val(Me.dcschemeid.BoundText)
    RsSavRec.Fields("Street").value = (txtstreetname.Text)
    RsSavRec.Fields("Northlength").value = val(txtnorthlength.Text)
    RsSavRec.Fields("Southlength").value = val(txtSouthlength.Text)
    RsSavRec.Fields("Eastlength").value = val(txteastlength.Text)
    RsSavRec.Fields("Westlength").value = val((txtWestlength.Text))
    RsSavRec.Fields("NorthlengthStr").value = (TxtPriceHadW.Text)
    RsSavRec.Fields("SouthlengthStr").value = (TxtPriceSomW.Text)
    RsSavRec.Fields("EastlengthStr").value = (TxteastWriiten.Text)
    RsSavRec.Fields("WestlengthStr").value = (TxtwestWriiten.Text)
    RsSavRec.Fields("SharMetre").value = val(TxtSharMetre.Text)
    RsSavRec.Fields("BanckName").value = TxtBanckName.Text
    RsSavRec.Fields("RecorDate").value = XPDtbTrans.value
    RsSavRec.Fields("BankID").value = val(Me.DcbBank.BoundText)
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("InvManager").value = val(Me.DcboEmpName.BoundText)
    RsSavRec.Fields("TypePrstg").value = val(Me.DcbTypePrstg.ListIndex)
    RsSavRec.Fields("PercenValue").value = val(Me.TxtPercenValue.Text)
    RsSavRec.Fields("InviseOrder").value = val(Me.TxtInviseOrder.Text)
    RsSavRec.Fields("InviseNo").value = val(Me.TxtInviseNo.Text)
    RsSavRec.Fields("InviseValue").value = val(Me.TxtInviseValue.Text)
    RsSavRec.Fields("SharesCount").value = val(Me.TxtSharesCount.Text)
    RsSavRec.Fields("SharesValue").value = val(Me.TxtSharesValue.Text)
    RsSavRec.Fields("Typ").value = val(Me.DcbType.ListIndex)
    RsSavRec.Fields("Googlemap").value = (txtgooglemap.Text)
    RsSavRec.Fields("Remarks").value = Me.TxtRemarks
    RsSavRec.Fields("Area").value = val(TxtArea.Text)
    RsSavRec.Fields("No_planned").value = TxtNo_planned.Text
    RsSavRec.Fields("MeterValue").value = val(TxtMeterValue.Text)
    RsSavRec.Fields("OwnerID").value = val(Me.dcsupplier.BoundText)
    RsSavRec.Fields("Identityof").value = Me.TxtIdentityof.Text
    RsSavRec.Fields("Address").value = Me.TxtAddress.Text
    RsSavRec.Fields("PropertyDeed").value = Me.TxtPropertyDeed.Text
    RsSavRec.Fields("PaymentType").value = val(Me.DcbPaymentType.ListIndex)
    RsSavRec.Fields("PaymentNo").value = val(Me.TxtPaymentNo.Text)
    RsSavRec.Fields("PeriodType").value = val(Me.DcbPeriodsID.ListIndex)
    RsSavRec.Fields("Period").value = val(Me.TxtPeriod.Text)
    RsSavRec.Fields("FristDate").value = FristDate.value
    RsSavRec.Fields("TotalInviseValue").value = val(TxtTotalInviseValue.Text)
    RsSavRec.Fields("TotalArea").value = val(TxtTotalArea.Text)
    RsSavRec.Fields("SharesValue2").value = val(TxtSharesValue2.Text)
    RsSavRec.Fields("SharesCount2").value = val(TxtSharesCount2.Text)
    RsSavRec.Fields("SharesMeters").value = val(TxtSharesMeters.Text)
    RsSavRec.Fields("TotalValue").value = val(TxtTotalValue.Text)
    RsSavRec.Fields("LandOwnedID").value = val(Me.DcbLand.BoundText)
    RsSavRec.Fields("Ownership").value = val(Me.Dcbownership.ListIndex)
    RsSavRec.Fields("AddValue").value = val(TxtAdd.Text)
    RsSavRec.Fields("DevelpoValue").value = val(TxtDevelpoValue.Text)
    RsSavRec.Fields("LandValue").value = val(TxtLandValue.Text)
    
        If Opt(0).value = True Then
    RsSavRec.Fields("Typepartial").value = 0
    ElseIf Opt(1).value = True Then
    RsSavRec.Fields("Typepartial").value = 1
    ElseIf Opt(2).value = True Then
    RsSavRec.Fields("Typepartial").value = 2
    Else
    RsSavRec.Fields("Typepartial").value = Null
    
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ''/////
    RsSavRec.Fields("Description").value = Me.TxtDescription.Text
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    
    sql = "update Tblinvestment set FlagActive=1 where ID=" & val(TxtInviseNo.Text) & ""
    Cn.Execute sql
  
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblActivateInvestmentDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim str2 As String
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("PaymentNo"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("ActInvID").value = val(Me.TxtSerial1.Text)
                RsDevsub("DatePayment").value = IIf((.TextMatrix(i, .ColIndex("DatePayment"))) = "", Null, .TextMatrix(i, .ColIndex("DatePayment")))
                RsDevsub("PaymentValue").value = IIf((.TextMatrix(i, .ColIndex("PaymentValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("PaymentValue"))))
                RsDevsub("Remrk").value = IIf((.TextMatrix(i, .ColIndex("Remrk"))) = "", Null, .TextMatrix(i, .ColIndex("Remrk")))
                RsDevsub("PaymentNo").value = IIf((.TextMatrix(i, .ColIndex("PaymentNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("PaymentNo"))))
       RsDevsub.update
      End If
     Next i
    End With
'''///////////////
If val(Me.Dcbownership.ListIndex) = 0 Then
  FilrecordLand
Else
  sql = "update TblBuyLanReEst set FlagActive=1,ActivID=" & val(TxtSerial1.Text) & " where ID=" & val(DcbLand.BoundText) & ""
    Cn.Execute sql
    
  End If
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
                FiLLTXT
                TxtModFlg = "R"
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
Function GetTOtalArea() As Double
Dim Rs5 As ADODB.Recordset
Dim sql As String
Set Rs5 = New ADODB.Recordset
sql = " SELECT     SUM(Area) AS SumTotalArea, InviseOrder"
sql = sql & " From dbo.TblActivateInvestment"
sql = sql & " Where (InviseOrder = " & val(TxtInviseOrder.Text) & ") AND (ID <> " & val(TxtSerial1.Text) & ")AND (LandOwnedID <>" & val(DcbLand.BoundText) & ")"
sql = sql & " GROUP BY InviseOrder"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
GetTOtalArea = IIf(IsNull(Rs5("SumTotalArea").value), 0, Rs5("SumTotalArea").value)
Else
GetTOtalArea = 0
End If
End Function
Function GetDiArea() As Double
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     SUM(ShareInvsCount) AS SumShareInvsCount, OrderInvse"
sql = sql & " From dbo.TblIPOSharer"
sql = sql & " Where (OrderInvse = " & val(TxtInviseOrder.Text) & ")"
sql = sql & " GROUP BY OrderInvse"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetDiArea = IIf(IsNull(Rs4("SumShareInvsCount").value), 0, Rs4("SumShareInvsCount").value)
Else
GetDiArea = 0
End If
End Function
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim i As Integer
        Me.TxtNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)

Me.TxtNoteSerial.Text = IIf(IsNull(RsSavRec("NoteSerial").value), "", RsSavRec("NoteSerial").value)
TxtRemarks2.Text = IIf(IsNull(RsSavRec.Fields("Remarks2").value), "", RsSavRec.Fields("Remarks2").value)

    TxtBanckName.Text = IIf(IsNull(RsSavRec.Fields("BanckName").value), "", RsSavRec.Fields("BanckName").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecorDate").value), Date, RsSavRec.Fields("RecorDate").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    Me.DcbBank.BoundText = IIf(IsNull(RsSavRec.Fields("BankID").value), "", RsSavRec.Fields("BankID").value)
    Me.DcboEmpName.BoundText = IIf(IsNull(RsSavRec.Fields("InvManager").value), "", RsSavRec.Fields("InvManager").value)
    Me.DcbTypePrstg.ListIndex = IIf(IsNull(RsSavRec.Fields("TypePrstg").value), -1, RsSavRec.Fields("TypePrstg").value)
    Me.TxtPercenValue.Text = IIf(IsNull(RsSavRec.Fields("PercenValue").value), 0, RsSavRec.Fields("PercenValue").value) ': ProgressBar1.value = 90
    Me.TxtInviseOrder.Text = IIf(IsNull(RsSavRec.Fields("InviseOrder").value), 0, RsSavRec.Fields("InviseOrder").value) ': ProgressBar1.value = 100
    RetriveInvist val(Me.TxtInviseOrder.Text)
    Me.txtgooglemap.Text = IIf(IsNull(RsSavRec.Fields("Googlemap").value), 0, RsSavRec.Fields("Googlemap").value)
    Me.TxtInviseNo.Text = IIf(IsNull(RsSavRec.Fields("InviseNo").value), 0, RsSavRec.Fields("InviseNo").value) ': ProgressBar1.value = 10
    Me.TxtInviseValue.Text = IIf(IsNull(RsSavRec.Fields("InviseValue").value), 0, RsSavRec.Fields("InviseValue").value) ': ProgressBar1.value = 20
    Me.TxtSharesCount.Text = IIf(IsNull(RsSavRec.Fields("SharesCount").value), 0, RsSavRec.Fields("SharesCount").value) ': ProgressBar1.value = 30
    TxtSharesValue.Text = IIf(IsNull(RsSavRec.Fields("SharesValue").value), 0, RsSavRec.Fields("SharesValue").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value) ': ProgressBar1.value = 40
    Me.DcbType.ListIndex = IIf(IsNull(RsSavRec.Fields("Typ").value), -1, RsSavRec.Fields("Typ").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    TxtArea.Text = IIf(IsNull(RsSavRec.Fields("Area").value), "", RsSavRec.Fields("Area").value)
    TxtNo_planned.Text = IIf(IsNull(RsSavRec.Fields("No_planned").value), "", RsSavRec.Fields("No_planned").value)
    TxtMeterValue.Text = IIf(IsNull(RsSavRec.Fields("MeterValue").value), 0, RsSavRec.Fields("MeterValue").value)
    Me.dcsupplier.BoundText = IIf(IsNull(RsSavRec.Fields("OwnerID").value), "", RsSavRec.Fields("OwnerID").value)
    TxtIdentityof.Text = IIf(IsNull(RsSavRec.Fields("Identityof").value), "", RsSavRec.Fields("Identityof").value)
    TxtAddress.Text = IIf(IsNull(RsSavRec.Fields("Address").value), "", RsSavRec.Fields("Address").value)
    TxtPropertyDeed.Text = IIf(IsNull(RsSavRec.Fields("PropertyDeed").value), "", RsSavRec.Fields("PropertyDeed").value)
    Me.DcbPaymentType.ListIndex = IIf(IsNull(RsSavRec.Fields("PaymentType").value), -1, RsSavRec.Fields("PaymentType").value)
    TxtPaymentNo.Text = IIf(IsNull(RsSavRec.Fields("PaymentNo").value), 0, RsSavRec.Fields("PaymentNo").value)
    TxtPeriod.Text = IIf(IsNull(RsSavRec.Fields("Period").value), 0, RsSavRec.Fields("Period").value)
    Me.DcbPeriodsID.ListIndex = IIf(IsNull(RsSavRec.Fields("PeriodType").value), -1, RsSavRec.Fields("PeriodType").value)
    FristDate.value = IIf(IsNull(RsSavRec.Fields("FristDate").value), Date, RsSavRec.Fields("FristDate").value)
    TxtTotalInviseValue.Text = IIf(IsNull(RsSavRec.Fields("TotalInviseValue").value), 0, RsSavRec.Fields("TotalInviseValue").value)
    TxtTotalArea.Text = IIf(IsNull(RsSavRec.Fields("TotalArea").value), 0, RsSavRec.Fields("TotalArea").value)
    TxtSharesValue2.Text = IIf(IsNull(RsSavRec.Fields("SharesValue2").value), 0, RsSavRec.Fields("SharesValue2").value)
    TxtSharesCount2.Text = IIf(IsNull(RsSavRec.Fields("SharesCount2").value), 0, RsSavRec.Fields("SharesCount2").value)
    TxtSharesMeters.Text = IIf(IsNull(RsSavRec.Fields("SharesMeters").value), 0, RsSavRec.Fields("SharesMeters").value)
    TxtTotalValue.Text = IIf(IsNull(RsSavRec.Fields("TotalValue").value), 0, RsSavRec.Fields("TotalValue").value)
    Me.Dcbownership.ListIndex = IIf(IsNull(RsSavRec.Fields("Ownership").value), -1, RsSavRec.Fields("Ownership").value)
    Me.DcbLand.BoundText = IIf(IsNull(RsSavRec.Fields("LandOwnedID").value), 0, RsSavRec.Fields("LandOwnedID").value)
    ''///////////////
    Me.DcboCountryID2.BoundText = IIf(IsNull(RsSavRec.Fields("CountryID").value), 0, RsSavRec.Fields("CountryID").value)
    Me.DcboGovernmentID.BoundText = IIf(IsNull(RsSavRec.Fields("CityID").value), 0, RsSavRec.Fields("CityID").value)
    Me.DcboCityID.BoundText = IIf(IsNull(RsSavRec.Fields("HyID").value), 0, RsSavRec.Fields("HyID").value)
    Me.dcschemeid.BoundText = IIf(IsNull(RsSavRec.Fields("SchemeID").value), 0, RsSavRec.Fields("SchemeID").value)
    Me.txtstreetname.Text = IIf(IsNull(RsSavRec.Fields("Street").value), "", RsSavRec.Fields("Street").value)
    Me.TxtPriceHadW.Text = IIf(IsNull(RsSavRec.Fields("NorthlengthStr").value), "", RsSavRec.Fields("NorthlengthStr").value)
    Me.TxtPriceSomW.Text = IIf(IsNull(RsSavRec.Fields("SouthlengthStr").value), "", RsSavRec.Fields("EastlengthStr").value)
    Me.TxteastWriiten.Text = IIf(IsNull(RsSavRec.Fields("EastlengthStr").value), "", RsSavRec.Fields("SouthlengthStr").value)
    Me.TxtwestWriiten.Text = IIf(IsNull(RsSavRec.Fields("WestlengthStr").value), "", RsSavRec.Fields("WestlengthStr").value)
    Me.txtnorthlength.Text = IIf(IsNull(RsSavRec.Fields("Northlength").value), 0, RsSavRec.Fields("Northlength").value)
    Me.txtSouthlength.Text = IIf(IsNull(RsSavRec.Fields("Southlength").value), 0, RsSavRec.Fields("Southlength").value)
    Me.txteastlength.Text = IIf(IsNull(RsSavRec.Fields("Eastlength").value), 0, RsSavRec.Fields("Eastlength").value)
    Me.txtWestlength.Text = IIf(IsNull(RsSavRec.Fields("Westlength").value), 0, RsSavRec.Fields("Westlength").value)
    Me.TxtAdd.Text = IIf(IsNull(RsSavRec.Fields("AddValue").value), 0, RsSavRec.Fields("AddValue").value)
    Me.TxtDevelpoValue.Text = IIf(IsNull(RsSavRec.Fields("DevelpoValue").value), 0, RsSavRec.Fields("DevelpoValue").value)
    Me.TxtLandValue.Text = IIf(IsNull(RsSavRec.Fields("LandValue").value), 0, RsSavRec.Fields("LandValue").value)
    If Not (IsNull(RsSavRec.Fields("Typepartial").value)) Then
    If RsSavRec.Fields("Typepartial").value = 0 Then
    Opt(0).value = True
    ElseIf RsSavRec.Fields("Typepartial").value = 1 Then
    Opt(1).value = True
    ElseIf RsSavRec.Fields("Typepartial").value = 2 Then
    Opt(2).value = True
    End If
    End If
    Me.TxtSharMetre.Text = IIf(IsNull(RsSavRec.Fields("SharMetre").value), 0, RsSavRec.Fields("SharMetre").value)
    
    TxtDescription.Text = IIf(IsNull(RsSavRec.Fields("Description").value), "", RsSavRec.Fields("Description").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData
ErrTrap:

End Sub


Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
RelinGrid
End Sub

Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With GridInstallments
Select Case .ColKey(Col)
Case "PaymentNo"
Cancel = True
Case "DatePayment"
Cancel = True
Case "PaymentValue"
If Opt(2).value = True Then
Cancel = False
Else
Cancel = True
End If

End Select
End With

End Sub

Private Sub ISButton2_Click()
 If Opt(0).value = False And Opt(1).value = False And Opt(2).value = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
        Else
        MsgBox "Please Select Method Number of decimal"
        End If
        Exit Sub
        End If
        
If val(TxtTotalValue.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· «·ﬁÌ„… «·«Ã„«·Ì…"
Else
MsgBox "Please Enter Total Value"
End If
'salah TxtTotalValue.SetFocus
Exit Sub
End If
If val(TxtPaymentNo.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· ⁄œœ «·œ›⁄«  "
Else
MsgBox "Please Enter No of  Payments"
End If
TxtPaymentNo.SetFocus
Exit Sub
End If
If val(TxtPeriod.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈œŒ«· «·› —… »Ì‰ «·œ›⁄«  "
Else
MsgBox "Please Enter No of  Period"
End If
TxtPeriod.SetFocus
Exit Sub
End If
If val(DcbPeriodsID.ListIndex) = -1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«—    ‰Ê⁄ «·› —… "
Else
MsgBox "Please Enter Type of  Period"
End If
DcbPeriodsID.SetFocus
Exit Sub
End If
filgrid1
RelinGrid
End Sub
Sub filgrid1()
Dim i As Integer
 GridInstallments.Clear flexClearScrollable, flexClearEverything
GridInstallments.Rows = 1
With GridInstallments

.Rows = .Rows + val(TxtPaymentNo.Text)
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("PaymentValue")) = val(TxtTotalValue.Text) / val(TxtPaymentNo.Text)
.TextMatrix(i, .ColIndex("PaymentValue")) = Round(val(.TextMatrix(i, .ColIndex("PaymentValue"))), 2)
.TextMatrix(i, .ColIndex("PaymentNo")) = i
If i = 1 Then
.TextMatrix(i, .ColIndex("DatePayment")) = FristDate.value
TempDate.value = FristDate.value
Else
If val(Me.DcbPeriodsID.ListIndex) = 0 Then
TempDate.value = DateAdd("d", val(Me.TxtPeriod.Text), TempDate.value)
ElseIf val(Me.DcbPeriodsID.ListIndex) = 1 Then
TempDate.value = DateAdd("M", val(Me.TxtPeriod.Text), TempDate.value)
ElseIf val(Me.DcbPeriodsID.ListIndex) = 1 Then
TempDate.value = DateAdd("YYYY", val(Me.TxtPeriod.Text), TempDate.value)
End If
.TextMatrix(i, .ColIndex("DatePayment")) = TempDate.value
End If
.TextMatrix(i, .ColIndex("Remrk")) = TxtRemarks.Text

Next i
'.AutoSize 0, .Cols - 1, False
End With
End Sub

Private Sub ISButton3_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments TxtSerial1.Text, "170420165"
ErrTrap:
End Sub

Private Sub ISButton5_Click()
print_report
End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
RevenueAccount = get_account_code_branch(129, my_branch)
If val(TxtAdd) > 0 Then


     If RevenueAccount = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                 MsgBox "Not creation  branch", vbCritical
                
                End If
               Exit Sub
            Else

                If RevenueAccount = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «Ì—«œ«   Œ’Ì’ «—«÷Ì    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                    MsgBox "Is not specified sales account", vbCritical
                    End If
       Exit Sub
                End If
            End If
       
       
End If
         
                        If DcbType.ListIndex = 0 Then
            RootAccount1 = get_account_code_branch(113, my_branch)
        Else
        RootAccount1 = get_account_code_branch(114, my_branch)
        End If
        
            If RootAccount1 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                 MsgBox "Not creation  branch", vbCritical
                
                End If
               Exit Sub
            Else

                If RootAccount1 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·„»Ì⁄«     ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                    MsgBox "Is not specified sales account", vbCritical
                    End If
       Exit Sub
                End If
            End If
        
 
                                   If DcbType.ListIndex = 0 Then
            RootAccount2 = get_account_code_branch(117, my_branch)
        Else
        RootAccount2 = get_account_code_branch(118, my_branch)
        End If
        
            If RootAccount2 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                MsgBox "Not creation  branch", vbCritical
                End If
               Exit Sub
            Else

                If RootAccount2 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·«—»«Õ    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                    MsgBox "Did not specify the expense of profits", vbCritical
                  
                End If
                Exit Sub
                End If
            End If
                 
                 

                                   If DcbType.ListIndex = 0 Then
            RootAccount3 = get_account_code_branch(119, my_branch)
        Else
        RootAccount3 = get_account_code_branch(120, my_branch)
        End If
        
            If RootAccount3 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                MsgBox "Not creation  branch", vbCritical
                End If
               Exit Sub
            Else

                If RootAccount3 = "NO account" Then
                                 If SystemOptions.UserInterface = ArabicInterface Then
                                     MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·Œ”«∆—    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                                    Else
                              MsgBox " Did not specify account losses", vbCritical
                        End If
                        Exit Sub
                        
                End If
            End If
                 
                 
                 
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double
Dim i As Integer
Sm = 0
     With GridInstallments
      For i = .FixedRows To .Rows - 1
         If Opt(0).value = True And i = 1 Then
            .TextMatrix(i, .ColIndex("PaymentValue")) = val(.TextMatrix(i, .ColIndex("PaymentValue"))) + ((val(TxtTotalValue.Text) - val(lbl(20).Caption)))
            .TextMatrix(i, .ColIndex("PaymentValue")) = Round(val(.TextMatrix(i, .ColIndex("PaymentValue"))), 2)
            End If
             If Opt(1).value = True And i = (.Rows - 1) Then
            
            .TextMatrix(i, .ColIndex("PaymentValue")) = val(.TextMatrix(i, .ColIndex("PaymentValue"))) + ((val(TxtTotalValue.Text) - val(lbl(20).Caption)))
           .TextMatrix(i, .ColIndex("PaymentValue")) = Round(val(.TextMatrix(i, .ColIndex("PaymentValue"))), 2)
            End If
            
        Next i
      End With
      RelinGrid
      
With Me.GridInstallments
If .Rows > 1 Then

If Round(val(lbl(20).Caption), 2) <> Round(val(TxtTotalValue.Text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ·«‰  „Ã„Ê⁄ ﬁÌ„ «·œ›⁄«  ·«Ì”«ÊÌ  «·ﬁÌ„… «·«Ã„«·Ì…"
Else
MsgBox "Can not Save  The values of Payement not equal the Total Value"
End If
Exit Sub
End If
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ  Ì—ÃÏ  Ê“Ì⁄ «·ﬁÌ„… ⁄·Ï «·œ›⁄«   "
Else
MsgBox "Can not Save "
End If
Exit Sub
End If
End With

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
   If val(Me.Dcbownership.ListIndex) = 0 Then
   If TxtDescription.Text = "" Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "Ì—ÃÏ «œŒ«· «”„ «·«—÷ «Ê «·⁄ﬁ«—"
   Else
   MsgBox "Eneter Name of Land"
   End If
   TxtDescription.SetFocus
   Exit Sub
   End If
  End If
           If DcbTypePrstg.Text = "" And val(DcbTypePrstg.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— ‰Ê⁄ ⁄„Ê·… «·«œ«—…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Commission ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            DcbTypePrstg.SetFocus
            Exit Sub
     End If
              
   If DcbTypePrstg.ListIndex <> 2 Then
              If val(TxtPercenValue.Text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈œŒ«·  ﬁÌ„…  ⁄„Ê·… «·«œ«—…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Eneter Commission ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            TxtPercenValue.SetFocus
            Exit Sub
     End If
     End If
           If DcboEmpName.Text = "" And val(DcboEmpName.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— „œÌ— «·„”«Â„… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Manager ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            DcboEmpName.SetFocus
            Exit Sub
     End If
     
      If val(TxtInviseNo.Text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· —ﬁ„ «·„”«Â„…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
                  Else
            MsgBox "Please Enter Investment  No", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          End If
'          TxtInviseNo.SetFocus
            Exit Sub
     End If
       If dcsupplier.Text = "" And val(dcsupplier.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ≈Œ Ì«— «·„«·ﬂ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Owner ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
           dcsupplier.SetFocus
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
    StrRecID = new_id("TblActivateInvestment", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Private Sub ISButton8_Click()
FrmSearchinvestment.inde = 6
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim ID  As Integer
GetTblBuyLandRealEstate ID, Me.Text1.Text, 1
DcbLand.BoundText = ID
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text10.Text, EmpID
        dcsupplier.BoundText = EmpID
    End If

End Sub



Private Sub TxtAdd_Change()
Calculte
End Sub

Private Sub TxtAdd_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtAdd.Text, 0)
End Sub

Private Sub TxtArea_Change()
Calculte
End Sub

Private Sub TxtArea_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtArea.Text, 0)
End Sub

Private Sub TxtDevelpoValue_Change()
Calculte
End Sub

Private Sub TxtInviseNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtInviseNo.Text, 0)
End Sub

Private Sub TxtInviseOrder_Change()
Dim ShrTotal As Double
Dim ShrTota2 As Double
If Me.TxtModFlg.Text <> "R" Then

If val(Me.TxtInviseOrder.Text) <> 0 Then
GetInvestInformation val(Me.TxtInviseOrder.Text), , ShrTotal
TxtSharesCount2.Text = ShrTotal
ShrTota2 = GetDiArea()
'''''''
If CheciIPOBySal_SharCount(val(Me.TxtInviseOrder.Text)) = False Then
If ShrTota2 <> ShrTotal Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰  ›⁄Ì· Â–Â «·„”«Â„… «·« »⁄œ «·«‰ Â«¡ „‰ «·«ﬂ  «»"
Else
MsgBox "You can not activate the Investment until after the completion of the IPO"
End If
TxtInviseOrder.SetFocus
   clear_all Me

       GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
    DcbType.ListIndex = 0
    Dcbownership.ListIndex = 0
    
Exit Sub
End If
End If
'''''
RetriveIPO val(Me.TxtInviseOrder.Text)
RetriveInvist val(Me.TxtInviseOrder.Text)
TxtTotalArea.Text = val(TxtArea.Text) + GetTOtalArea()
 End If
End If
End Sub
Sub RetriveIPO(Optional OrderInvse As Double = 0)
If OrderInvse <> 0 Then
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = "Select * from TblIPO where OrderInvse=" & OrderInvse & ""
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
TxtSharesCount.Text = IIf(IsNull(Rs8("CountShare").value), "", Rs8("CountShare").value)
TxtSharesValue.Text = IIf(IsNull(Rs8("ShareValue").value), 0, Rs8("ShareValue").value)
TxtSharesValue.Text = Round(TxtSharesValue.Text, 2)
Else
TxtSharesValue.Text = 0
TxtSharesCount.Text = 0
End If
End If
End Sub
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
  sql = "SELECT    * from  TblActivateInvestmentDet"
  sql = sql + "  Where (ActInvID = " & val(TxtSerial1.Text) & ") "
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("DatePayment")) = IIf(IsNull(Rs1("DatePayment").value), "", Rs1("DatePayment").value)
                   .TextMatrix(i, .ColIndex("PaymentValue")) = IIf(IsNull(Rs1("PaymentValue").value), 0, Rs1("PaymentValue").value)
                   .TextMatrix(i, .ColIndex("PaymentNo")) = IIf(IsNull(Rs1("PaymentNo").value), 0, Rs1("PaymentNo").value)
                   .TextMatrix(i, .ColIndex("Remrk")) = IIf(IsNull(Rs1("Remrk").value), "", Rs1("Remrk").value)
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub
Sub RetriveInvist(Optional ID As Double = 0)
If ID <> 0 Then
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = "Select * from Tblinvestment where id=" & ID & ""
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
TxtInviseNo.Text = IIf(IsNull(Rs8("ID").value), "", Rs8("ID").value)
TxtInviseValue.Text = IIf(IsNull(Rs8("TotalInDe").value), 0, Rs8("TotalInDe").value)
DcbBank.BoundText = IIf(IsNull(Rs8("BankID").value), "", Rs8("BankID").value)
DcbType.ListIndex = IIf(IsNull(Rs8("Typ").value), -1, Rs8("Typ").value)
TxtBanckName.Text = IIf(IsNull(Rs8("BanckName").value), "", Rs8("BanckName").value)
tXTRootAccount = IIf(IsNull(Rs8("ParentAccount").value), "", Rs8("ParentAccount").value)
tXTNAME = IIf(IsNull(Rs8("NAME").value), "", Rs8("NAME").value)

Else
TxtBanckName.Text = ""
TxtInviseNo.Text = 0
TxtInviseValue.Text = 0
DcbBank.BoundText = 0
DcbType.ListIndex = -1
tXTRootAccount = ""
tXTNAME = ""
End If
End If
End Sub

Private Sub TxtInviseOrder_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearchinvestment.inde = 5
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End If
End Sub

Private Sub TxtInviseValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtInviseValue.Text, 0)
End Sub

Private Sub TxtLandValue_Change()
Calculte
End Sub

Private Sub TxtMeterValue_Change()
Calculte
End Sub
Sub Calculte()
If Me.TxtModFlg.Text <> "R" Then
Me.TxtLandValue.Text = (val(TxtArea.Text) * val(TxtMeterValue.Text))
TxtTotalValue.Text = val(Me.TxtLandValue.Text) + val(TxtAdd.Text) + val(Me.TxtDevelpoValue.Text)
TxtTotalValue.Text = Round(val(TxtTotalValue.Text), 2)
End If
End Sub


Private Sub TxtMeterValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtMeterValue.Text, 0)
End Sub

Private Sub TxtPaymentNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtPaymentNo.Text, 0)
End Sub

Private Sub TxtPeriod_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtPeriod.Text, 0)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
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
      If CheckExp(TxtInviseNo.Text) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… ·«‰Â« „— »ÿ… » ÿÊÌ— «·«—«÷Ì"
    Else
    MsgBox "Can not Delete This process linked to the development of land"
    End If
    Exit Sub
    End If
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
      Dim StrSQL As String
          
                               StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords

          
                    Dim Account_code1 As String
          Dim Account_code2 As String
          Dim Account_code3 As String
          Dim Account_code4 As String
          Dim Account_Code5 As String
          Dim Account_Code6  As String
          Dim ParentAccount  As String
          Dim ParentAccount1  As String
          Dim ParentAccountsub  As String
          Dim Account_Code7 As String
          
          
         
          Account_code3 = IIf(IsNull(RsSavRec("Account_Code3").value), "", RsSavRec("Account_Code3").value)
          Account_code4 = IIf(IsNull(RsSavRec("Account_Code4").value), "", RsSavRec("Account_Code4").value)
          Account_Code5 = IIf(IsNull(RsSavRec("Account_Code5").value), "", RsSavRec("Account_Code5").value)
          Account_Code6 = IIf(IsNull(RsSavRec("Account_Code6").value), "", RsSavRec("Account_Code6").value)
           Account_Code7 = IIf(IsNull(RsSavRec("Account_Code7").value), "", RsSavRec("Account_Code7").value)
          
           
          ParentAccount1 = IIf(IsNull(RsSavRec("ParentAccount1").value), "", RsSavRec("ParentAccount1").value)
           
           
           
           
 If ModAccounts.CheckDeleteAccount(Account_code3, True) = True _
 And ModAccounts.CheckDeleteAccount(Account_code4, True) = True _
 And ModAccounts.CheckDeleteAccount(Account_Code5, True) = True _
 And ModAccounts.CheckDeleteAccount(Account_Code6, True) = True _
  And ModAccounts.CheckDeleteAccount(Account_Code7, True) = True _
 Then
     

 Else
 If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·« Ì„ﬂ‰ «·Õ–› ÌÊÃœ Õ—ﬂ«  ⁄·Ì Õ”«» «·„”«Â„…", vbCritical
    Else
    MsgBox "You can not delete.There are movements at the expense of of contributions", vbCritical
    End If
                Exit Sub
                
 GoTo ErrTrap
  End If
 
           
           
 If ModAccounts.DeleteAccount(Account_code3, True) = True _
 And ModAccounts.DeleteAccount(Account_code4, True) = True _
 And ModAccounts.DeleteAccount(Account_Code5, True) = True _
 And ModAccounts.DeleteAccount(Account_Code6, True) = True _
 And ModAccounts.DeleteAccount(Account_Code7, True) = True _
  Then
                If ModAccounts.DeleteAccount(ParentAccount1) = True Then
                Else
                GoTo ErrTrap
                End If
   
                
  Else
 
  GoTo ErrTrap
  End If
 
Dim sql As String
 sql = "update Tblinvestment set FlagActive=null where ID=" & val(TxtInviseNo.Text) & ""
    Cn.Execute sql
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                  StrSQL = "Delete From TblActivateInvestmentDet Where ActInvID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                
                                    '  If ModAccounts.DeleteAccount(StrAccountCode, True) = True Then
                                      ' CuurentLogdata ("D")
                                          RsSavRec.delete
                                  '      Msg = " „  ⁄„·Ì… «·Õ–›."
                                  '      MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            
                                    'Else
                                    '   Exit Sub
                                    'End If
                                    
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
             
               '''''''''''''''''''''''''''''''
               StrSQL = "update TblBuyLanReEst set FlagActive=null where ID=" & val(DcbLand.BoundText) & ""
    Cn.Execute StrSQL
 StrSQL = "Delete From TblBuyLanReEst Where ActivID =" & val(TxtSerial1.Text) & ""
 Cn.Execute StrSQL, , adExecuteNoRecords
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If

     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
               LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
           ' StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
           ' StrMSG = "You can not delete the record"
           '† StrMSG = StrMSG & " Is related to with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           'Cn.Errors.Clear
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
       
        DcbType.Enabled = True
        
    ElseIf TxtModFlg.Text = "R" Then
     XPDtbTrans.Enabled = False
      DcbType.Enabled = False
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
DcbType.Enabled = False
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
    If CheckExp(TxtInviseNo.Text) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰  ⁄œÌ· Â–Â «·Õ—ﬂ… ·«‰Â« „— »ÿ… » ÿÊÌ— «·«—«÷Ì"
    Else
    MsgBox "Can not Update This process linked to the development of land"
    End If
    Exit Sub
    End If
        TxtModFlg = "E"
            GridInstallments.Rows = GridInstallments.Rows + 1
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

    TxtModFlg.Text = "N"
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
    DcbType.ListIndex = 0
    Dcbownership.ListIndex = 0
DcbTypePrstg.ListIndex = 2

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
  MySQL = "SELECT     dbo.TblActivateInvestment.ID, dbo.TblActivateInvestment.RecorDate, dbo.TblActivateInvestment.TypePrstg, dbo.TblActivateInvestment.PercenValue, "
  MySQL = MySQL & "                    dbo.TblActivateInvestment.InviseOrder, dbo.TblActivateInvestment.InviseNo, dbo.TblActivateInvestment.InviseValue, dbo.TblActivateInvestment.SharesCount,"
  MySQL = MySQL & "                     dbo.TblActivateInvestment.SharesValue, dbo.TblActivateInvestment.Typ, dbo.TblActivateInvestment.Description, dbo.TblActivateInvestment.Remarks,"
  MySQL = MySQL & "                    dbo.TblActivateInvestment.Area, dbo.TblActivateInvestment.No_planned, dbo.TblActivateInvestment.MeterValue, dbo.TblActivateInvestment.Identityof,"
  MySQL = MySQL & "                    dbo.TblActivateInvestment.Address, dbo.TblActivateInvestment.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
  MySQL = MySQL & "                    dbo.TblActivateInvestment.InvManager, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
  MySQL = MySQL & "                    dbo.TblActivateInvestment.OwnerID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode,"
  MySQL = MySQL & "                    dbo.TblActivateInvestment.TotalValue, dbo.TblActivateInvestment.SharesMeters, dbo.TblActivateInvestment.SharesCount2, dbo.TblActivateInvestment.TotalArea,"
  MySQL = MySQL & "                    dbo.TblActivateInvestment.TotalInviseValue, dbo.TblActivateInvestment.FristDate, dbo.TblActivateInvestment.PeriodType, dbo.TblActivateInvestment.SharesValue2,"
  MySQL = MySQL & "                    dbo.TblActivateInvestment.Period, dbo.TblActivateInvestment.PaymentNo, dbo.TblActivateInvestment.BankID, dbo.BanksData.BankName, dbo.BanksData.BankNamee,"
  MySQL = MySQL & "                     dbo.TblActivateInvestment.PaymentType, dbo.TblActivateInvestment.PropertyDeed, dbo.TblActivateInvestmentDet.DatePayment,"
  MySQL = MySQL & "                    dbo.TblActivateInvestmentDet.PaymentValue, dbo.TblActivateInvestmentDet.PaymentNo AS DetPaymentNo, dbo.TblActivateInvestmentDet.Remrk"
  MySQL = MySQL & "  FROM         dbo.TblActivateInvestment LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblActivateInvestmentDet ON dbo.TblActivateInvestment.ID = dbo.TblActivateInvestmentDet.ActInvID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.BanksData ON dbo.TblActivateInvestment.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCustemers ON dbo.TblActivateInvestment.OwnerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblEmployee ON dbo.TblActivateInvestment.InvManager = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.TblActivateInvestment.BranchID = dbo.TblBranchesData.branch_id"
  MySQL = MySQL & " Where (dbo.TblActivateInvestment.id =" & val(TxtSerial1.Text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepActiveInvesment.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepActiveInvesment.rpt"
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
Sub GetInformationBuyLand(Optional ID As Integer = 0)
Dim sql As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset

sql = "Select * from TblBuyLanReEst where id=" & ID & ""
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
'dcschemeid.BoundText = IIf(IsNull(Rs6("SchemeID").value), 0, Rs6("SchemeID").value)
DcboCountryID2.BoundText = IIf(IsNull(Rs6("CountryID").value), 0, Rs6("CountryID").value)
DcboGovernmentID.BoundText = IIf(IsNull(Rs6("CityID").value), 0, Rs6("CityID").value)
DcboCityID.BoundText = IIf(IsNull(Rs6("HyID").value), 0, Rs6("HyID").value)
txtstreetname.Text = IIf(IsNull(Rs6("Street").value), "", Rs6("Street").value)
TxtAddress.Text = IIf(IsNull(Rs6("DesLocation").value), "", Rs6("DesLocation").value)
txtnorthlength.Text = IIf(IsNull(Rs6("Northlength").value), "", Rs6("Northlength").value)
TxtPriceHadW.Text = IIf(IsNull(Rs6("NorthlengthStr").value), "", Rs6("NorthlengthStr").value)
txtSouthlength.Text = IIf(IsNull(Rs6("Southlength").value), "", Rs6("Southlength").value)
TxtPriceSomW.Text = IIf(IsNull(Rs6("SouthlengthStr").value), "", Rs6("SouthlengthStr").value)
txteastlength.Text = IIf(IsNull(Rs6("Eastlength").value), "", Rs6("Eastlength").value)
TxteastWriiten.Text = IIf(IsNull(Rs6("EastlengthStr").value), "", Rs6("EastlengthStr").value)
txtWestlength.Text = IIf(IsNull(Rs6("Westlength").value), "", Rs6("Westlength").value)
TxtwestWriiten.Text = IIf(IsNull(Rs6("WestlengthStr").value), "", Rs6("WestlengthStr").value)
TxtPropertyDeed.Text = IIf(IsNull(Rs6("TitledeedNo").value), "", Rs6("TitledeedNo").value)
TxtNo_planned.Text = IIf(IsNull(Rs6("SchemName").value), "", Rs6("SchemName").value)
TxtMeterValue.Text = IIf(IsNull(Rs6("MeterPrice").value), 0, Rs6("MeterPrice").value)
TxtArea.Text = IIf(IsNull(Rs6("Area").value), 0, Rs6("Area").value)
dcsupplier.BoundText = IIf(IsNull(Rs6("OwnerID").value), 0, Rs6("OwnerID").value)
txtgooglemap.Text = IIf(IsNull(Rs6("Googlemap").value), "", Rs6("Googlemap").value)
Else
txtgooglemap.Text = ""
TxtDescription.Text = ""
dcsupplier.BoundText = 0
TxtArea.Text = ""
TxtMeterValue.Text = 0
TxtNo_planned.Text = ""
dcschemeid.BoundText = 0
DcboCountryID2.BoundText = 0
DcboGovernmentID.BoundText = 0
DcboCityID.BoundText = 0
txtstreetname.Text = ""
TxtAddress.Text = ""
txtnorthlength.Text = ""
TxtPriceHadW.Text = ""
txtSouthlength.Text = ""
TxtPriceSomW.Text = ""
txteastlength.Text = ""
TxteastWriiten.Text = ""
txtWestlength.Text = ""
TxtwestWriiten.Text = ""
TxtPropertyDeed.Text = ""
End If

End Sub

Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
   ''''''''''''''''''''////
      Label7.Caption = "Numbers"
   Label19.Caption = "Writing"
   Frame7.Caption = "Border"
   Label3.Caption = "North"
   Label23.Caption = "North"
   Label5.Caption = "South"
   Label21.Caption = "South"
   Label4.Caption = "East"
   Label22.Caption = "East"
   Label6.Caption = "West"
   Label20.Caption = "West"

   Label1(0).Caption = "Date"
   Label1(3).Caption = "Country"
   Label1(4).Caption = "City"
   Label1(6).Caption = "Street"
   Label1(5).Caption = "District"
   Label1(15).Caption = "Scheme"

   ''/////////////
    Me.Caption = "Investment  "
    Label1(0).Caption = "Land Owned"
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(4).Caption = "Trans ID"
    Me.lbl(2).Caption = "Date"
   lbl(22).Caption = "Investment No"
    lbl(1).Caption = "Inve. No"
    Me.lbl(7).Caption = "Branch"
    lbl(3).Caption = "Investment Value"
    lbl(15).Caption = "Manager"
    lbl(0).Caption = "Admi Commission"
    lbl(27).Caption = "Count Share"
    lbl(5).Caption = "Count Share"
    lbl(6).Caption = "Share Value"
    lbl(16).Caption = "Type"
    Label1(8).Caption = "No.Payments"
    Label1(9).Caption = "First Date"
    lbl(9).Caption = "Description"
    Label1(11).Caption = "Period"
    lbl(10).Caption = "No.Planned"
    ISButton2.Caption = "Add"
    lbl(28).Caption = "No Shares Meteres"
    lbl(21).Caption = "Remarks"
    lbl(12).Caption = "Total Area"
    lbl(25).Caption = "Total Area"
    lbl(19).Caption = "Total Value"
    lbl(24).Caption = "Total Value"
    lbl(13).Caption = "Meter Value"
    lbl(26).Caption = "Meter Value"
    Label1(1).Caption = "Owner"
    lbl(14).Caption = "Identification No "
    lbl(17).Caption = "Address"
    lbl(23).Caption = "Property Deed"
    lbl(11).Caption = "Payment"
    lbl(18).Caption = "Bank"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    '''''''''''''' next
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
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
    Frame6.Caption = "Data of Payments"
  With Me.GridInstallments
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("PaymentNo")) = "Payment No"
  .TextMatrix(0, .ColIndex("DatePayment")) = "Date Payment"
  .TextMatrix(0, .ColIndex("PaymentValue")) = "Payment Value"
  .TextMatrix(0, .ColIndex("Remrk")) = "Remarks"
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
lbl(20).Caption = 0
With Me.GridInstallments
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("PaymentNo"))) <> 0 Then
Counter = Counter + 1
.TextMatrix(i, .ColIndex("Ser")) = Counter
summation = summation + val(.TextMatrix(i, .ColIndex("PaymentValue")))
End If
Next i
lbl(20).Caption = summation

End With
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblActivateInvestment"
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


Private Sub TxtSharesCount_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtSharesCount.Text) <> 0 Then
TxtSharMetre.Text = val(TxtTotalArea.Text) / val(TxtSharesCount.Text)
TxtSharMetre.Text = Round(val(TxtSharMetre.Text), 2)
End If
End If
End Sub

Private Sub TxtSharesCount_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtSharesCount.Text, 0)
End Sub

Private Sub TxtSharesCount2_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtTotalArea.Text) <> 0 Then
TxtSharesMeters.Text = Round(val(TxtSharesCount2.Text) / val(TxtTotalArea.Text), 2)
TxtSharesMeters.Text = Round(TxtSharesMeters.Text, 2)
End If
End If
End Sub

Private Sub TxtSharesMeters_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtSharesMeters.Text, 0)
End Sub

Private Sub TxtSharesValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtSharesValue.Text, 0)
End Sub

Private Sub TxtSharesValue2_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtSharesValue2.Text, 0)
End Sub

Private Sub TxtTotalArea_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtSharesCount.Text) <> 0 Then
TxtSharMetre.Text = val(TxtTotalArea.Text) / val(TxtSharesCount.Text)
TxtSharMetre.Text = Round(val(TxtSharMetre.Text), 2)
End If
End If
End Sub
Function GetDevelopValue(Optional LandID As Double) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     SUM(dbo.TblExpensesInvesmentDet.Valu *dbo.TblExpensesInvesmentDet.TypTrns) AS SumValu, dbo.TblExpensesInvesment.LandID"
sql = sql & " FROM         dbo.TblExpensesInvesmentDet RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblExpensesInvesment ON dbo.TblExpensesInvesmentDet.ExpInvID = dbo.TblExpensesInvesment.ID"
sql = sql & " Where (dbo.TblExpensesInvesment.TypDiv = 1) And (dbo.TblExpensesInvesment.LandID = " & LandID & ")"
sql = sql & " GROUP BY dbo.TblExpensesInvesment.LandID"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetDevelopValue = IIf(IsNull(Rs3("SumValu").value), 0, Rs3("SumValu").value)
Else
GetDevelopValue = 0
End If
End Function
Private Sub TxtTotalArea_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtTotalArea.Text, 0)
End Sub

Private Sub TxtTotalInviseValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtTotalInviseValue.Text, 0)
End Sub



Private Sub TxtTotalValue_Change()
If Me.TxtModFlg.Text <> " R" Then
TxtTotalInviseValue.Text = val(Me.TxtTotalValue.Text)
End If
End Sub

Private Sub TxtTotalValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtTotalValue.Text, 0)
End Sub

Private Sub XPDtbTrans_Change()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.Text = ""
End If
End Sub
