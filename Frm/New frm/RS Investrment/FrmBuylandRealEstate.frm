VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmBuylandRealEstate 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14760
   Icon            =   "FrmBuylandRealEstate.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   14760
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmBuylandRealEstate.frx":6852
      Left            =   15480
      List            =   "FrmBuylandRealEstate.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   63
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
      TabIndex        =   57
      Top             =   0
      Width           =   14745
      Begin VB.TextBox txtopening_balance_voucher_id 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   151
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox tXTRootAccount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   58
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
         ButtonImage     =   "FrmBuylandRealEstate.frx":687B
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   59
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
         ButtonImage     =   "FrmBuylandRealEstate.frx":6C15
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   60
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
         ButtonImage     =   "FrmBuylandRealEstate.frx":6FAF
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   61
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
         ButtonImage     =   "FrmBuylandRealEstate.frx":7349
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‘—«¡ «·«—«÷Ì /«·⁄ﬁ«—« "
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
         TabIndex        =   62
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmBuylandRealEstate.frx":76E3
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
      TabIndex        =   52
      Top             =   720
      Width           =   14715
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·œ›⁄« "
         ForeColor       =   &H00C00000&
         Height          =   3255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   5280
         Width           =   14535
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√Ê· ﬁ”ÿ"
            Height          =   252
            Index           =   0
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   600
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
            TabIndex        =   129
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœÊÌ"
            Height          =   252
            Index           =   2
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   555
            Left            =   240
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   240
            Width           =   4815
         End
         Begin VB.ComboBox DcbPeriodsID 
            Height          =   315
            ItemData        =   "FrmBuylandRealEstate.frx":8AE8
            Left            =   6240
            List            =   "FrmBuylandRealEstate.frx":8AEA
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   240
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
            Left            =   7440
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
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
            Left            =   12240
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   1065
         End
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   1875
            Left            =   120
            TabIndex        =   93
            Top             =   960
            Width           =   14325
            _cx             =   25268
            _cy             =   3307
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
            FormatString    =   $"FrmBuylandRealEstate.frx":8AEC
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
            Height          =   270
            Left            =   9600
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   94961667
            CurrentDate     =   41640
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   315
            Left            =   6240
            TabIndex        =   37
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   600
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
            ButtonImage     =   "FrmBuylandRealEstate.frx":8BB1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSComCtl2.DTPicker TempDate 
            Height          =   270
            Left            =   -960
            TabIndex        =   99
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
            Format          =   94961667
            CurrentDate     =   41640
         End
         Begin ImpulseButton.ISButton ISButton6 
            Height          =   330
            Left            =   12600
            TabIndex        =   149
            ToolTipText     =   "Õ–› «·ﬂ·"
            Top             =   2880
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› ’›"
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
            ButtonImage     =   "FrmBuylandRealEstate.frx":F413
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   330
            Left            =   10680
            TabIndex        =   150
            ToolTipText     =   "Õ–› «·ﬂ·"
            Top             =   2880
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
            ButtonImage     =   "FrmBuylandRealEstate.frx":15C75
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… Ã»— «·ﬂ”Ê—"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   37
            Left            =   12600
            TabIndex        =   130
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Index           =   6
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   2880
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
            TabIndex        =   103
            Top             =   2880
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   21
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·› —Â »Ì‰ «·œ›⁄« "
            Height          =   285
            Index           =   11
            Left            =   8040
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «Ê· œ›⁄Â"
            Height          =   285
            Index           =   9
            Left            =   10920
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·œ›⁄« "
            Height          =   285
            Index           =   8
            Left            =   13440
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   240
            Width           =   930
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   4935
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   14655
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Height          =   4815
            Left            =   -120
            TabIndex        =   84
            Top             =   0
            Width           =   14775
            Begin VB.TextBox TxtRemarks2 
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
               Left            =   6150
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   157
               Top             =   2760
               Width           =   7185
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·—’Ìœ «·√›  «ÕÏ «·Ã«—Ì"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   915
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   140
               Top             =   1920
               Width           =   5745
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ‰"
                  Height          =   255
                  Index           =   0
                  Left            =   2670
                  RightToLeft     =   -1  'True
                  TabIndex        =   143
                  Top             =   210
                  Width           =   765
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ«∆‰"
                  Height          =   255
                  Index           =   1
                  Left            =   1800
                  RightToLeft     =   -1  'True
                  TabIndex        =   142
                  Top             =   210
                  Width           =   765
               End
               Begin VB.OptionButton OptType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€Ì— „Õœœ"
                  Height          =   255
                  Index           =   2
                  Left            =   810
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   1005
               End
               Begin VB.TextBox TxtOpenBalance 
                  Alignment       =   2  'Center
                  Height          =   285
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   510
                  Width           =   1365
               End
               Begin MSComCtl2.DTPicker Dtp 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   144
                  Top             =   510
                  Width           =   1380
                  _ExtentX        =   2434
                  _ExtentY        =   503
                  _Version        =   393216
                  Enabled         =   0   'False
                  CalendarBackColor=   12648447
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   94961667
                  CurrentDate     =   38718
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   285
                  Index           =   15
                  Left            =   1320
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   450
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «·—’Ìœ "
                  Height          =   255
                  Index           =   14
                  Left            =   3900
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   540
                  Width           =   1275
               End
            End
            Begin VB.ComboBox DcbPaymentType 
               Height          =   315
               ItemData        =   "FrmBuylandRealEstate.frx":1C4D7
               Left            =   6120
               List            =   "FrmBuylandRealEstate.frx":1C4D9
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox TxtTotalValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   600
               Width           =   1935
            End
            Begin VB.TextBox TxtFullCode 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   11640
               RightToLeft     =   -1  'True
               TabIndex        =   0
               Top             =   240
               Width           =   1695
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
               Left            =   7590
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   2400
               Width           =   5745
            End
            Begin VB.CommandButton Command3 
               Caption         =   "⁄—÷"
               Height          =   315
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   2400
               Width           =   1455
            End
            Begin VB.TextBox TxtSchemName 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   8760
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               Top             =   960
               Width           =   1455
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÕœÊœ"
               ForeColor       =   &H00C00000&
               Height          =   975
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   3600
               Width           =   14535
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
                  Left            =   9600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
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
                  Left            =   6600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   240
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
                  Left            =   9600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
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
                  Left            =   6600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
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
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘„«·"
                  Height          =   255
                  Left            =   11400
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
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
                  TabIndex        =   126
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ã‰Ê»"
                  Height          =   255
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€—»"
                  Height          =   255
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   240
                  Width           =   855
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÕœÊœ «—ﬁ«„"
                  Height          =   255
                  Left            =   12480
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   360
                  Width           =   1215
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÕœÊœ ﬂ «»Â"
                  Height          =   255
                  Left            =   12480
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "€—»"
                  Height          =   255
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ã‰Ê»"
                  Height          =   255
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
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
                  TabIndex        =   119
                  Top             =   600
                  Width           =   855
               End
               Begin VB.Label Label23 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘„«·"
                  Height          =   255
                  Left            =   11400
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   720
                  Width           =   855
               End
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Œ·’… »«·«—«÷Ì"
               ForeColor       =   &H00C00000&
               Height          =   735
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   3000
               Width           =   14535
               Begin VB.TextBox TxtUnit 
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
                  Left            =   10560
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.TextBox TxtBlock 
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
                  Left            =   7080
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.TextBox TxtPart 
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
                  Left            =   3600
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   240
                  Width           =   1665
               End
               Begin VB.TextBox TxtStreet 
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
                  TabIndex        =   23
                  Top             =   240
                  Width           =   2145
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Ã«œ…"
                  Height          =   255
                  Left            =   12240
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·»·Êﬂ"
                  Height          =   255
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label17 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «··ÊÕ…"
                  Height          =   255
                  Left            =   5280
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·‘Ê«—⁄"
                  Height          =   255
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin VB.TextBox txtlocation 
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
               Left            =   6150
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Top             =   2040
               Width           =   7185
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
               Height          =   435
               Left            =   240
               MaxLength       =   50
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               Tag             =   "«œŒ· «”„ «·‘«—⁄"
               Top             =   1440
               Width           =   4095
            End
            Begin VB.TextBox TxtNameE 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   2
               Top             =   240
               Width           =   4095
            End
            Begin VB.TextBox TxtPropertyDeed 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   11640
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   600
               Width           =   1695
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
               Left            =   9150
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   1320
               Width           =   1065
            End
            Begin VB.TextBox TxtMeterValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TxtArea 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   8760
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   600
               Width           =   1455
            End
            Begin VB.TextBox TxtNo_planned 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   11640
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox NameTxt 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   1
               Top             =   240
               Width           =   4095
            End
            Begin MSDataListLib.DataCombo dcsupplier 
               Height          =   315
               Left            =   6120
               TabIndex        =   10
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   1320
               Width           =   2955
               _ExtentX        =   5212
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin Dynamic_Byte.NourHijriCal DatePropertyDeed 
               Height          =   315
               Left            =   11640
               TabIndex        =   8
               Top             =   1320
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
            End
            Begin MSDataListLib.DataCombo DcboCountryID2 
               Height          =   315
               Left            =   11640
               TabIndex        =   13
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
               Left            =   8760
               TabIndex        =   14
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
               Left            =   6120
               TabIndex        =   15
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
               Top             =   1680
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcschemeid 
               Height          =   315
               Left            =   7200
               TabIndex        =   42
               Tag             =   " "
               Top             =   1920
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„·«ÕŸ« "
               Height          =   285
               Index           =   7
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   2760
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   405
               Index           =   17
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   960
               Width           =   4275
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   285
               Index           =   16
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   147
               Top             =   960
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ"
               Height          =   285
               Index           =   3
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Ê’› «·„Êﬁ⁄"
               Height          =   285
               Index           =   33
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   2400
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·„Œÿÿ"
               Height          =   285
               Index           =   15
               Left            =   10200
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   960
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Ê’› «·„Êﬁ⁄"
               Height          =   285
               Index           =   29
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   2040
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·‘«—⁄"
               Height          =   285
               Index           =   6
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   1560
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·ÕÌ"
               Height          =   285
               Index           =   5
               Left            =   7440
               RightToLeft     =   -1  'True
               TabIndex        =   108
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
               Left            =   10200
               RightToLeft     =   -1  'True
               TabIndex        =   107
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
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   1680
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ ’ﬂ «·„·ﬂÌ…"
               Height          =   285
               Index           =   0
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   1320
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ê’› ≈‰Ã·Ì“Ì"
               Height          =   285
               Index           =   0
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "’ﬂ «·„·ﬂÌ…"
               Height          =   285
               Index           =   23
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   600
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   285
               Index           =   19
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   600
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·œ›⁄"
               Height          =   285
               Index           =   11
               Left            =   7665
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   990
               Width           =   1050
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «·„«·ﬂ"
               Height          =   285
               Index           =   1
               Left            =   10200
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   1320
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "”⁄— «·„ —"
               Height          =   285
               Index           =   13
               Left            =   7440
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   600
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„”«Õ… "
               Height          =   285
               Index           =   12
               Left            =   10200
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   600
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·„Œÿÿ"
               Height          =   285
               Index           =   10
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   960
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ê’› ⁄—»Ì"
               Height          =   285
               Index           =   1
               Left            =   10200
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   240
               Width           =   1515
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         TabIndex        =   53
         Top             =   0
         Width           =   14655
         Begin VB.TextBox TxtBillNo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   240
            Width           =   1215
         End
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   134
            Top             =   240
            Width           =   975
            _Version        =   786432
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "‘—«¡"
            ForeColor       =   8388608
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   12840
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   10680
            TabIndex        =   40
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   94961665
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmBuylandRealEstate.frx":1C4DB
            Height          =   315
            Left            =   6600
            TabIndex        =   41
            Top             =   240
            Width           =   3375
            _ExtentX        =   5953
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
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   136
            Top             =   240
            Width           =   975
            _Version        =   786432
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "„—œÊœ« "
            ForeColor       =   8388608
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   139
            Top             =   240
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«—’œ… «›  «ÕÌ…"
            ForeColor       =   8388608
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï ›« Ê—… —ﬁ„"
            Height          =   285
            Index           =   9
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   9480
            TabIndex        =   83
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ "
            Height          =   285
            Index           =   4
            Left            =   13800
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   11970
            TabIndex        =   54
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
      TabIndex        =   51
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
      TabIndex        =   50
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   65
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
      TabIndex        =   66
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
      Height          =   1425
      Left            =   0
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   9120
      Width           =   14715
      _cx             =   25956
      _cy             =   2514
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
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   152
         Top             =   0
         Width           =   4605
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   240
            Width           =   1095
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
            TabIndex        =   156
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   69
         Top             =   0
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   73
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
            TabIndex        =   72
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
            TabIndex        =   71
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
            TabIndex        =   70
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   68
         Top             =   600
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   44
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
            ButtonImage     =   "FrmBuylandRealEstate.frx":1C4F0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8880
            TabIndex        =   46
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
            ButtonImage     =   "FrmBuylandRealEstate.frx":22D52
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   45
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
            ButtonImage     =   "FrmBuylandRealEstate.frx":230EC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   47
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
            ButtonImage     =   "FrmBuylandRealEstate.frx":2994E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   48
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
            ButtonImage     =   "FrmBuylandRealEstate.frx":29CE8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   49
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
            ButtonImage     =   "FrmBuylandRealEstate.frx":2A282
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   3960
            TabIndex        =   81
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
            ButtonImage     =   "FrmBuylandRealEstate.frx":2A61C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1920
            TabIndex        =   82
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
            ButtonImage     =   "FrmBuylandRealEstate.frx":30E7E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   10440
         TabIndex        =   74
         Top             =   240
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   570
         Left            =   120
         TabIndex        =   78
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
         Left            =   4080
         TabIndex        =   133
         ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
         Top             =   120
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
         ButtonImage     =   "FrmBuylandRealEstate.frx":31218
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
         TabIndex        =   75
         Top             =   240
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
            Picture         =   "FrmBuylandRealEstate.frx":37A7A
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuylandRealEstate.frx":37E14
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuylandRealEstate.frx":381AE
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuylandRealEstate.frx":38548
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuylandRealEstate.frx":388E2
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuylandRealEstate.frx":38C7C
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuylandRealEstate.frx":39016
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuylandRealEstate.frx":395B0
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   76
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
      ButtonImage     =   "FrmBuylandRealEstate.frx":3994A
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   79
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
      ButtonImage     =   "FrmBuylandRealEstate.frx":401AC
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   80
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
      ButtonImage     =   "FrmBuylandRealEstate.frx":46A0E
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
      TabIndex        =   77
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmBuylandRealEstate"
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
 Dim RecID As String
 Dim Account_Code_dynamic As String
 Dim Account_Code_dynamic1 As String
 Dim II As Long
 Public LonRow As Double
Public LngCol As Double

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
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
      '  Msg = EleHeader.Caption & " ??? " & txtID & " EEC??I" & Date
  If TxtRemarks2.Text = "" Then
 If Rd(0).value = True Then
Msg = " ‘—«¡ «·«—«÷Ì »—ﬁ„ " & TxtFullcode & "··«—÷" & NameTxt.Text & "  ··„«·ﬂ  " & dcsupplier.Text
 
ElseIf Rd(1).value = True Then
Msg = " „—œÊœ«  ‘—«¡ «·«—«÷Ì »—ﬁ„ " & TxtFullcode & "··«—÷" & NameTxt.Text & "  ··„«·ﬂ  " & dcsupplier.Text
End If
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
If Rd(0).value = True Then
DebitAcc = RsSavRec("Account_Code").value
CreditAcc = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcsupplier.BoundText))
ElseIf Rd(1).value = True Then
DebitAcc = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcsupplier.BoundText))
CreditAcc = RsSavRec("Account_Code").value

End If


line_no = 1
 
    BranchID = val(Dcbranch.BoundText)
    
           
                 If ModAccounts.AddNewDev(LngDevID, line_no, DebitAcc, val(TxtTotalValue), 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                'CreditAcc
                If ModAccounts.AddNewDev(LngDevID, line_no, CreditAcc, val(TxtTotalValue), 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
            
                
                
                
             
     
     
'     Next i
     
'     End With
           
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
If Rd(0).value = True Then
des = " ‘—«¡ «·«—«÷Ì »—ﬁ„ " & TxtFullcode & "··«—÷" & NameTxt.Text & "  ··„«·ﬂ  " & dcsupplier.Text
 notytype = 9001
ElseIf Rd(1).value = True Then
des = " „—œÊœ«  ‘—«¡ «·«—«÷Ì »—ﬁ„ " & TxtFullcode & "··«—÷" & NameTxt.Text & "  ··„«·ﬂ  " & dcsupplier.Text
notytype = 9002
End If

Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim Sql As String
tablename = "TblBuyLanReEst"
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
                                                                 Sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                Sql = Sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   Sql = Sql & " where NoteID=" & val(TxtNoteID.Text)
                                                                   Cn.Execute Sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TxtNoteID.Text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent
 

     End If

End Function

Private Sub Dcbranch_Change()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.Text = ""
End If
End Sub

Private Sub Dcbranch_Click(Area As Integer)
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.Text = ""
End If
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

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
 

  If SystemOptions.UserInterface = ArabicInterface Then
    With DcbPaymentType
    .Clear
    .AddItem "‰ﬁœÌ"
    .AddItem "«Ã·"
    End With
       With DcbPeriodsID
    .Clear
    .AddItem "ÌÊ„"
    .AddItem "‘Â—"
    .AddItem "”‰…"
    End With
 Else
     With DcbPeriodsID
    .Clear
    .AddItem "Day"
    .AddItem "Month"
    .AddItem "Year"
    End With
    With DcbPaymentType
    .Clear
    .AddItem "Cash"
    .AddItem "Credit"
    End With
End If
LoadDataCombos
loadcombo
    conection = "select * from TblBuyLanReEst WHERE     (NewLand IS NULL) order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
 
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetUsers Me.DCboUserName
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
   FiLLTXT
ErrTrap:
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
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  
  
  '  On Error GoTo ErrTrap
    Dim Sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                 StrSQL = "Delete From TblBuyLanReEstInsta Where ByLanID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
      StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords

              
              
              End If
            
            
              If Me.TxtModFlg.Text = "N" Then
                RsSavRec("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, Trim$(Me.NameTxt.Text), True, False, Me.TxtNameE.Text, , , , , , , , , , 1, 1, 1, 0, 0)
                 
            Else

                        If Not IsNull(RsSavRec("Account_Code").value) Then
                            ModAccounts.EditAccount RsSavRec("Account_Code").value, Me.NameTxt.Text, Me.TxtNameE.Text, , , , , , , , , 1, 1, 1, 0, 0, , , , True
                        End If
            End If
            
    RsSavRec.Fields("Name").value = NameTxt.Text
    RsSavRec.Fields("NameE").value = TxtNameE.Text
    RsSavRec.Fields("FullCode").value = TxtFullcode.Text
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("Remarks2").value = TxtRemarks2.Text
    '''////
    RsSavRec.Fields("CountryID").value = val(Me.DcboCountryID2.BoundText)
    RsSavRec.Fields("CityID").value = val(Me.DcboGovernmentID.BoundText)
    RsSavRec.Fields("HyID").value = val(Me.DcboCityID.BoundText)
    RsSavRec.Fields("SchemeID").value = val(Me.dcschemeid.BoundText)
    RsSavRec.Fields("DesLocation").value = (txtlocation.Text)
    RsSavRec.Fields("Street").value = (txtstreetname.Text)
    RsSavRec.Fields("DatePropertyDeed").value = DatePropertyDeed.value
    RsSavRec.Fields("Unit").value = (TxtUnit.Text)
    RsSavRec.Fields("Block").value = (TxtBlock.Text)
    RsSavRec.Fields("PlateNo").value = (TxtPart.Text)
    RsSavRec.Fields("StreetNo").value = (TxtStreet.Text)
    RsSavRec.Fields("Northlength").value = val(TxtNorthLength.Text)
    RsSavRec.Fields("Southlength").value = val(TxtSouthLength.Text)
    RsSavRec.Fields("Eastlength").value = val(TxtEastLength.Text)
    RsSavRec.Fields("Westlength").value = val((txtWestlength.Text))
    RsSavRec.Fields("NorthlengthStr").value = (TxtPriceHadW.Text)
    RsSavRec.Fields("SouthlengthStr").value = (TxtPriceSomW.Text)
    RsSavRec.Fields("EastlengthStr").value = (TxteastWriiten.Text)
    RsSavRec.Fields("WestlengthStr").value = (TxtwestWriiten.Text)
    
 ''''//////////////////////
    RsSavRec.Fields("Area").value = val(TxtArea.Text)
    RsSavRec.Fields("No_planned").value = TxtNo_planned.Text
    RsSavRec.Fields("MeterPrice").value = val(TxtMeterValue.Text)
    RsSavRec.Fields("OwnerID").value = val(Me.dcsupplier.BoundText)
    RsSavRec.Fields("Total").value = val(Me.TxtTotalValue.Text)
    RsSavRec.Fields("TitledeedNo").value = Me.TxtPropertyDeed.Text
    RsSavRec.Fields("PayType").value = val(Me.DcbPaymentType.ListIndex)
    RsSavRec.Fields("InstalNo").value = val(Me.TxtPaymentNo.Text)
    RsSavRec.Fields("PeriodType").value = val(Me.DcbPeriodsID.ListIndex)
    RsSavRec.Fields("Period").value = val(Me.TxtPeriod.Text)
    RsSavRec.Fields("FristDate").value = FristDate.value
    RsSavRec.Fields("Remarks").value = TxtRemarks.Text
    RsSavRec.Fields("SchemName").value = TxtSchemName.Text
    RsSavRec.Fields("Googlemap").value = txtgooglemap.Text
    If Rd(1).value = True Then
    RsSavRec.Fields("BillNo").value = val(TxtBillNo.Text)
    RsSavRec.Fields("FlgReturn").value = 1
    RsSavRec.Fields("BuySal").value = 1
    ElseIf Rd(2).value = True Then
    RsSavRec.Fields("BuySal").value = 2
    End If
    If OptType(0).value = True Then
    RsSavRec.Fields("Debt_Credit").value = 0
    ElseIf OptType(1).value = True Then
    RsSavRec.Fields("Debt_Credit").value = 1
     ElseIf OptType(2).value = True Then
    RsSavRec.Fields("Debt_Credit").value = 2
    End If
    RsSavRec.Fields("OpenBalance").value = val(TxtOpenBalance.Text)
    RsSavRec.Fields("OpenDate").value = Dtp.value
    
    If Opt(0).value = True Then
    RsSavRec.Fields("Typepartial").value = 0
    ElseIf Opt(1).value = True Then
    RsSavRec.Fields("Typepartial").value = 1
    ElseIf Opt(2).value = True Then
    RsSavRec.Fields("Typepartial").value = 2
    Else
    RsSavRec.Fields("Typepartial").value = Null
    
    End If
    
    If Rd(1).value = True Then
    Sql = "Update TblBuyLanReEst set FlgReturn=1 where id =" & val(TxtBillNo.Text) & " "
     Cn.Execute Sql
   End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ''/////
 
  RsSavRec("OpenBalanceDate").value = Me.Dtp.value
        If Me.OptType(2).value = True Then
            RsSavRec("OpenBalance").value = 0
            RsSavRec("OpenBalanceType").value = Null
        ElseIf Me.OptType(0).value = True Then
            RsSavRec("OpenBalance").value = val(Me.TxtOpenBalance.Text)
            RsSavRec("OpenBalanceType").value = 0
        ElseIf Me.OptType(1).value = True Then
            RsSavRec("OpenBalance").value = val(Me.TxtOpenBalance.Text)
            RsSavRec("OpenBalanceType").value = 1
        End If
          
       If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
       
       '     If val(Me.txtopening_balance_voucher_id.text) = 0 Then
                txtopening_balance_voucher_id.Text = get_opening_balance_voucher_id
               RsSavRec("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.Text)
       '     End If '
        End If '
'********************************
    
    
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    
    '********************************
    If Rd(2).value = True Then '«›  «ÕÌ
     Dim StrDes As String

        If SystemOptions.UserInterface = ArabicInterface Then
            StrDes = "—’Ìœ «› «ÕÌ ··«—÷ —ﬁ„" & Trim(Me.TxtFullcode.Text) & "    " & Me.NameTxt
        Else
            StrDes = " Opening Balance For: " & Trim(Me.TxtFullcode.Text) & "    " & TxtNameE.Text
        End If
        
        If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
            If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        
                Dim LngDevID As Long
                Dim LngOpenID As Long
            
                 LngOpenID = 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
           
                If Me.OptType(0).value = True Then
       
                    If ModAccounts.AddNewDev(LngDevID, 1, RsSavRec("Account_Code").value, val(Me.TxtOpenBalance.Text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                    If ModAccounts.AddNewDev(LngDevID, 2, Account_Code_dynamic1, val(Me.TxtOpenBalance.Text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                
                 ElseIf Me.OptType(1).value = True Then
            
                  
                    If ModAccounts.AddNewDev(LngDevID, 1, Account_Code_dynamic1, val(Me.TxtOpenBalance.Text), 0, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                 
                    If ModAccounts.AddNewDev(LngDevID, 2, RsSavRec("Account_Code").value, val(Me.TxtOpenBalance.Text), 1, StrDes, LngOpenID, , , , Me.Dtp.value, , , , , , , , , , , , , , True, val(txtopening_balance_voucher_id.Text), , , , val(Dcbranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If

                 
            End If
        End If
   ElseIf Rd(0).value = True Then '„‘ —Ì« 
   createVoucher
   ElseIf Rd(1).value = True Then '„—œÊœ« 
   createVoucher
   End If

'********************************

''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblBuyLanReEstInsta Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim str2 As String
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("PaymentNo"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("ByLanID").value = val(Me.TxtSerial1.Text)
                RsDevsub("InstalDate").value = IIf((.TextMatrix(i, .ColIndex("DatePayment"))) = "", Null, .TextMatrix(i, .ColIndex("DatePayment")))
                RsDevsub("Val").value = IIf((.TextMatrix(i, .ColIndex("PaymentValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("PaymentValue"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remrk"))) = "", Null, .TextMatrix(i, .ColIndex("Remrk")))
                RsDevsub("InstalNo").value = IIf((.TextMatrix(i, .ColIndex("PaymentNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("PaymentNo"))))
       RsDevsub.update
      End If
     Next i
    End With
'''///////////////
  
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & Chr(13)
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
 Private Sub OptType_Click(Index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.Text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.Text)
End Sub

Private Sub TxtOpenBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOpenBalance.Text, 0)
End Sub

Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim i As Integer
        Me.TxtNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)

Me.TxtNoteSerial.Text = IIf(IsNull(RsSavRec("NoteSerial").value), "", RsSavRec("NoteSerial").value)
Me.TxtRemarks2.Text = IIf(IsNull(RsSavRec("Remarks2").value), "", RsSavRec("Remarks2").value)
    NameTxt.Text = IIf(IsNull(RsSavRec.Fields("Name").value), "", RsSavRec.Fields("Name").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    Me.TxtFullcode.Text = IIf(IsNull(RsSavRec.Fields("FullCode").value), "", RsSavRec.Fields("FullCode").value)
    Me.TxtNameE.Text = IIf(IsNull(RsSavRec.Fields("NameE").value), "", RsSavRec.Fields("NameE").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value) ': ProgressBar1.value = 40
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    TxtArea.Text = IIf(IsNull(RsSavRec.Fields("Area").value), "", RsSavRec.Fields("Area").value)
    TxtTotalValue.Text = IIf(IsNull(RsSavRec.Fields("Total").value), "", RsSavRec.Fields("Total").value)
    TxtNo_planned.Text = IIf(IsNull(RsSavRec.Fields("No_planned").value), "", RsSavRec.Fields("No_planned").value)
    TxtMeterValue.Text = IIf(IsNull(RsSavRec.Fields("MeterPrice").value), 0, RsSavRec.Fields("MeterPrice").value)
    Me.dcsupplier.BoundText = IIf(IsNull(RsSavRec.Fields("OwnerID").value), "", RsSavRec.Fields("OwnerID").value)
    TxtPropertyDeed.Text = IIf(IsNull(RsSavRec.Fields("TitledeedNo").value), "", RsSavRec.Fields("TitledeedNo").value)
    Me.DcbPaymentType.ListIndex = IIf(IsNull(RsSavRec.Fields("PayType").value), -1, RsSavRec.Fields("PayType").value)
    TxtPaymentNo.Text = IIf(IsNull(RsSavRec.Fields("InstalNo").value), 0, RsSavRec.Fields("InstalNo").value)
    TxtPeriod.Text = IIf(IsNull(RsSavRec.Fields("Period").value), 0, RsSavRec.Fields("Period").value)
    Me.DcbPeriodsID.ListIndex = IIf(IsNull(RsSavRec.Fields("PeriodType").value), -1, RsSavRec.Fields("PeriodType").value)
    FristDate.value = IIf(IsNull(RsSavRec.Fields("FristDate").value), Date, RsSavRec.Fields("FristDate").value)
    '''//////////////
    Me.DcboCountryID2.BoundText = IIf(IsNull(RsSavRec.Fields("CountryID").value), 0, RsSavRec.Fields("CountryID").value)
    Me.DcboGovernmentID.BoundText = IIf(IsNull(RsSavRec.Fields("CityID").value), 0, RsSavRec.Fields("CityID").value)
    Me.DcboCityID.BoundText = IIf(IsNull(RsSavRec.Fields("HyID").value), 0, RsSavRec.Fields("HyID").value)
  '  Me.dcschemeid.BoundText = IIf(IsNull(RsSavRec.Fields("SchemeID").value), 0, RsSavRec.Fields("SchemeID").value)
    Me.txtlocation.Text = IIf(IsNull(RsSavRec.Fields("DesLocation").value), "", RsSavRec.Fields("DesLocation").value)
    Me.txtstreetname.Text = IIf(IsNull(RsSavRec.Fields("Street").value), "", RsSavRec.Fields("Street").value)
    DatePropertyDeed.value = IIf(IsNull(RsSavRec.Fields("DatePropertyDeed").value), ToHijriDate(Date), RsSavRec.Fields("DatePropertyDeed").value)
    Me.TxtUnit.Text = IIf(IsNull(RsSavRec.Fields("Unit").value), "", RsSavRec.Fields("Unit").value)
    Me.TxtPart.Text = IIf(IsNull(RsSavRec.Fields("PlateNo").value), "", RsSavRec.Fields("PlateNo").value)
    Me.TxtStreet.Text = IIf(IsNull(RsSavRec.Fields("StreetNo").value), "", RsSavRec.Fields("StreetNo").value)
    Me.TxtPriceHadW.Text = IIf(IsNull(RsSavRec.Fields("NorthlengthStr").value), "", RsSavRec.Fields("NorthlengthStr").value)
    Me.TxtPriceSomW.Text = IIf(IsNull(RsSavRec.Fields("SouthlengthStr").value), "", RsSavRec.Fields("EastlengthStr").value)
    Me.TxteastWriiten.Text = IIf(IsNull(RsSavRec.Fields("EastlengthStr").value), "", RsSavRec.Fields("EastlengthStr").value)
    Me.TxtwestWriiten.Text = IIf(IsNull(RsSavRec.Fields("WestlengthStr").value), "", RsSavRec.Fields("WestlengthStr").value)
    Me.TxtNorthLength.Text = IIf(IsNull(RsSavRec.Fields("Northlength").value), 0, RsSavRec.Fields("Northlength").value)
    Me.TxtSouthLength.Text = IIf(IsNull(RsSavRec.Fields("Southlength").value), 0, RsSavRec.Fields("Southlength").value)
    Me.TxtEastLength.Text = IIf(IsNull(RsSavRec.Fields("Eastlength").value), 0, RsSavRec.Fields("Eastlength").value)
    Me.txtWestlength.Text = IIf(IsNull(RsSavRec.Fields("Westlength").value), 0, RsSavRec.Fields("Westlength").value)
    Me.TxtSchemName.Text = IIf(IsNull(RsSavRec.Fields("SchemName").value), "", RsSavRec.Fields("SchemName").value)
    Me.txtgooglemap.Text = IIf(IsNull(RsSavRec.Fields("Googlemap").value), "", RsSavRec.Fields("Googlemap").value)
    txtopening_balance_voucher_id.Text = IIf(IsNull(RsSavRec("opening_balance_voucher_id").value), "", RsSavRec("opening_balance_voucher_id").value)


    If Not (IsNull(RsSavRec.Fields("Typepartial").value)) Then
    If RsSavRec.Fields("Typepartial").value = 0 Then
    Opt(0).value = True
    ElseIf RsSavRec.Fields("Typepartial").value = 1 Then
    Opt(1).value = True
    ElseIf RsSavRec.Fields("Typepartial").value = 2 Then
    Opt(2).value = True
    End If
    End If
    If Me.TxtModFlg.Text = "R" Then
    If Not (IsNull(RsSavRec.Fields("BuySal").value)) Then
    If RsSavRec.Fields("BuySal").value = 1 Then
    Rd(1).value = True
    ElseIf RsSavRec.Fields("BuySal").value = 2 Then
    Rd(2).value = True
    Else
    Rd(0).value = True
    End If
    Else
    Rd(0).value = True
    End If
    TxtBillNo.Text = IIf(IsNull(RsSavRec.Fields("BillNo").value), "", RsSavRec.Fields("BillNo").value)
    End If
    TxtOpenBalance.Text = IIf(IsNull(RsSavRec.Fields("OpenBalance").value), 0, RsSavRec.Fields("OpenBalance").value)
    Dtp.value = IIf(IsNull(RsSavRec.Fields("OpenDate").value), Date, RsSavRec.Fields("OpenDate").value)
    If Not (IsNull(RsSavRec.Fields("Debt_Credit").value)) Then
    If RsSavRec.Fields("Debt_Credit").value = 0 Then
    OptType(0).value = True
    ElseIf RsSavRec.Fields("Debt_Credit").value = 1 Then
    OptType(1).value = True
     ElseIf RsSavRec.Fields("Debt_Credit").value = 2 Then
    OptType(2).value = True
    End If
    End If

    
    If Not (IsNull(RsSavRec("OpenBalanceDate").value)) Then
        Me.Dtp.value = RsSavRec("OpenBalanceDate").value
         
    Else
    
        Me.Dtp.value = Date
        Me.Dtp.Enabled = False
    End If

    If Not IsNull(RsSavRec("OpenBalanceType").value) Then
        Me.TxtOpenBalance.Text = IIf(IsNull(RsSavRec("OpenBalance")), "", (RsSavRec("OpenBalance")))

        If RsSavRec("OpenBalanceType").value = 0 Then
            OptType(0).value = True
            OptType_Click 0
        ElseIf RsSavRec("OpenBalanceType").value = 1 Then
            OptType(1).value = True
            OptType_Click 1
        End If
    
    Else
        Me.TxtOpenBalance.Text = 0
        Me.OptType(2).value = True
        OptType_Click 2
    End If



    ''//////////
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData
RelinGrid
ErrTrap:
End Sub


Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
RelinGrid
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

Private Sub DcboCityID_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
loadcombo
End If
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
TxtTotalValue.SetFocus
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
.TextMatrix(i, .ColIndex("PaymentValue")) = (val(TxtTotalValue.Text) / val(TxtPaymentNo.Text))
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
            
ShowAttachments TxtSerial1.Text, "170420164"
ErrTrap:
End Sub

Private Sub ISButton4_Click()
If Me.TxtModFlg.Text <> "R" Then
        If CheActiveLand(val(TxtSerial1.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ…  „  ›⁄Ì·Â«"
Else
MsgBox "Can Not delete This Process Active"
End If
Exit Sub
End If
If Rd(0).value = True Then
If CheReturnLand(val(TxtSerial1.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ…  „ ⁄„· „—œÊœ ·Â« "
Else
MsgBox "Can Not delete this process Return "
End If
Exit Sub
End If
End If

GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
  End If
End Sub

Private Sub ISButton5_Click()
print_report
End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
    Dim FirstPeriodDateInthisYear As Date
    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

    Me.Dtp.value = FirstPeriodDateInthisYear
    
   ' On Error GoTo ErrTrap
         Dim Total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double
Dim i As Integer
                    Account_Code_dynamic1 = get_account_code_branch(109, my_branch)
        
                    If Account_Code_dynamic1 = "NO branch" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» Ê”Ìÿ «›  «ÕÌ ··«—«÷Ì", vbCritical
                        GoTo ErrTrap
                    Else

                        If Account_Code_dynamic1 = "NO account" Then
                            MsgBox "·„ Ì „  ÕœÌœ Õ”«» Ê”Ìÿ «›  «ÕÌ ··«—«÷Ì", vbCritical
                            GoTo ErrTrap
                     
                        End If
                    End If
                    
          Account_Code_dynamic = get_account_code_branch(127, my_branch)
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» ··«—«÷Ì    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
       
                End If
            End If
            
Sm = 0
     With GridInstallments
      For i = .FixedRows To .Rows - 1
         If Opt(0).value = True And i = 1 Then
            .TextMatrix(i, .ColIndex("PaymentValue")) = val(.TextMatrix(i, .ColIndex("PaymentValue"))) + ((val(TxtTotalValue.Text) - val(lbl(6).Caption)))
            .TextMatrix(i, .ColIndex("PaymentValue")) = Round(val(.TextMatrix(i, .ColIndex("PaymentValue"))), 2)
            End If
             If Opt(1).value = True And i = (.Rows - 1) Then
            
            .TextMatrix(i, .ColIndex("PaymentValue")) = val(.TextMatrix(i, .ColIndex("PaymentValue"))) + ((val(TxtTotalValue.Text) - val(lbl(6).Caption)))
           .TextMatrix(i, .ColIndex("PaymentValue")) = Round(val(.TextMatrix(i, .ColIndex("PaymentValue"))), 2)
            End If
            
        Next i
      End With
      RelinGrid
With Me.GridInstallments
If .Rows > 1 Then
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("PaymentNo"))) <> 0 Then
Sm = Sm + val(.TextMatrix(i, .ColIndex("PaymentValue")))
End If
Next i
Total = val(TxtTotalValue.Text)
Total = Round(Total, 2)
Sm = Round(Sm, 2)
If Sm <> Total Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ ·«‰  „Ã„Ê⁄ ﬁÌ„ «·œ›⁄«  ·«Ì”«ÊÌ  «·ﬁÌ„… «·«Ã„«·Ì…"
Else
MsgBox "Can not Save  The values of Payement not equal the Total Value"
End If
Exit Sub
End If
End If
End With
If Rd(1).value = True Then
If val(Me.TxtBillNo.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ «·›« Ê—…"
Else
MsgBox "Please Enter No of Bill"
End If
TxtBillNo.SetFocus
Exit Sub
End If
End If
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
           If TxtFullcode.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡≈œŒ«· «·ﬂÊœ  ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
           Else
            MsgBox "Please Eneter Code ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            TxtFullcode.SetFocus
          Exit Sub
        End If
     If val(TxtTotalValue.Text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...«·—Ã«¡ «œŒ«· «·ﬁÌ„… «·«Ã„«·Ì…"
     Else
     MsgBox "Please Eneter Total Value"
     End If
'     TxtTotalValue.SetFocus
     Exit Sub
     End If
      If NameTxt.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«·  «·Ê’› ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
                  Else
            MsgBox "Please Enter Description  ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          End If
          NameTxt.SetFocus
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
          StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
    
    
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
Function CheActiveLand(Optional ID As Double = 0) As Boolean
Dim Sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Sql = "select * from TblActivateInvestment where LandOwnedID =" & ID & " "
Rs4.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CheActiveLand = True
Else
CheActiveLand = False
End If
End Function
Function CheReturnLand(Optional ID As Double = 0) As Boolean
Dim Sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Sql = "select * from TblBuyLanReEst where ID =" & ID & " and FlgReturn=1 "
Rs4.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CheReturnLand = True
Else
CheReturnLand = False
End If
End Function
Function CheOpenBalnLand(Optional ID As Double = 0) As Boolean
Dim Sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Sql = "select * from TblBuyLanReEst where ID =" & ID & " and BuySal=2 "
Rs4.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CheOpenBalnLand = True
Else
CheOpenBalnLand = False
End If
End Function

' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblBuyLanReEst", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub


Private Sub ISButton6_Click()
If Me.TxtModFlg.Text <> "R" Then
        If CheActiveLand(val(TxtSerial1.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ…  „  ›⁄Ì·Â«"
Else
MsgBox "Can Not delete This Process Active"
End If
Exit Sub
End If
If Rd(0).value = True Then
If CheReturnLand(val(TxtSerial1.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ…  „ ⁄„· „—œÊœ ·Â« "
Else
MsgBox "Can Not delete this process Return "
End If
Exit Sub
End If
End If

With GridInstallments
If .Rows < 2 Then
Exit Sub
Else
.RemoveItem .Row
End If
End With
End If
End Sub

Private Sub ISButton8_Click()
FrmSearchinvestment.inde = 7
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End Sub

Private Sub NameTxt_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub



Private Sub RadioButton1_Click()

End Sub

Private Sub Rd_Click(Index As Integer)
Fra.Visible = False
Frame2.Enabled = True
Frame6.Enabled = True
TxtBillNo.Visible = False
lbl(9).Visible = False
If Rd(1).value = True Then
Frame2.Enabled = False
Frame6.Enabled = False
TxtBillNo.Visible = True
lbl(9).Visible = True
ElseIf Rd(2).value = True Then
Fra.Visible = True
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text10.Text, EmpID
        dcsupplier.BoundText = EmpID
    End If

End Sub


 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim Sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
  Sql = "SELECT    * from  TblBuyLanReEstInsta"
  Sql = Sql + "  Where (ByLanID = " & val(TxtSerial1.Text) & ") "
  Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("DatePayment")) = IIf(IsNull(Rs1("InstalDate").value), Date, Rs1("InstalDate").value)
                   .TextMatrix(i, .ColIndex("PaymentValue")) = IIf(IsNull(Rs1("Val").value), 0, Rs1("Val").value)
                   .TextMatrix(i, .ColIndex("PaymentNo")) = IIf(IsNull(Rs1("InstalNo").value), 0, Rs1("InstalNo").value)
                   .TextMatrix(i, .ColIndex("Remrk")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub



Private Sub TxtArea_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtTotalValue.Text = val(Me.TxtArea.Text) * val(TxtMeterValue.Text)
TxtTotalValue = Round(val(TxtTotalValue.Text), 2)
End If
End Sub

Private Sub TxtBillNo_Change()

If Me.TxtModFlg.Text <> "R" Then
If val(TxtBillNo.Text) <> 0 Then
If CheActiveLand(val(TxtBillNo.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ ⁄„· „—œÊœ«  Â–Â «·Õ—ﬂ…  „  ›⁄Ì·Â«"
Else
MsgBox "Can Not Return This Process Active"
End If
Exit Sub
End If
If CheReturnLand(val(TxtBillNo.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ ⁄„· „—œÊœ«  Â–Â «·Õ—ﬂ…  „ ⁄„· „—œÊœ ·Â« ”«»ﬁ«"
Else
MsgBox "Can Not Return "
End If
Exit Sub
End If

If CheOpenBalnLand(val(TxtBillNo.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ ⁄„· „—œÊœ«  Â–Â «·Õ—ﬂ…«—’œ… «›  «ÕÌ…"
Else
MsgBox "Can Not Return this process Balances opening "
End If
Exit Sub
End If
    RsSavRec.find "ID=" & val(TxtBillNo.Text), , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
      Else
End If
End If
End If
End Sub

Private Sub TxtArea_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtArea.Text, 0)
End Sub

Private Sub TxtBillNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtBillNo.Text, 0)
End Sub

Private Sub TxtBillNo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearchinvestment.inde = 71
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End If
End Sub

Private Sub TxtEastLength_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtEastLength.Text, 0)
End Sub

Private Sub TxtMeterValue_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtTotalValue.Text = val(Me.TxtArea.Text) * val(TxtMeterValue.Text)
TxtTotalValue = Round(val(TxtTotalValue.Text), 2)
End If
End Sub

Private Sub TxtMeterValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtMeterValue.Text, 0)
End Sub

Private Sub TxtNameE_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub TxtNorthLength_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtNorthLength.Text, 0)
End Sub

Private Sub TxtPaymentNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtPaymentNo.Text, 0)
End Sub

Private Sub TxtPeriod_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtPeriod.Text, 0)
End Sub



' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecID, , adSearchForward, 1
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
    Dim Sql As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double
        If CheActiveLand(val(TxtSerial1.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ…  „  ›⁄Ì·Â«"
Else
MsgBox "Can Not delete This Process Active"
End If
Exit Sub
End If
If Rd(0).value = True Then
If CheReturnLand(val(TxtSerial1.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ…  „ ⁄„· „—œÊœ ·Â« "
Else
MsgBox "Can Not delete this process Return "
End If
Exit Sub
End If
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
     
                          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords

    Sql = "Update TblBuyLanReEst set FlgReturn=null where id =" & val(TxtBillNo.Text) & " "
Cn.Execute Sql


                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                  StrSQL = "Delete From TblBuyLanReEstInsta Where ByLanID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                
                                  
  StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
       Dim StrAccountCode As String
                StrAccountCode = RsSavRec("Account_Code").value
     
            
           StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            
        If ModAccounts.DeleteAccount(StrAccountCode, True) = True Then
                                        RsSavRec.delete
                                        Msg = "   „ «·Õ–›  ."
                                        MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            
                                    Else
                                        GoTo ErrTrap
                                    End If
                     
                     
                                              
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
                   RecID As String)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
    If CheActiveLand(val(TxtSerial1.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·Õ—ﬂ…  „  ›⁄Ì·Â«"
Else
MsgBox "Can Not Edite This Process Active"
End If
Exit Sub
End If
If Rd(0).value = True Then
If CheReturnLand(val(TxtSerial1.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·Õ—ﬂ…  „ ⁄„· „—œÊœ ·Â« "
Else
MsgBox "Can Not edite this process Return "
End If
Exit Sub
End If
End If



        TxtModFlg = "E"
            'GridInstallments.Rows = GridInstallments.Rows + 1
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
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & Chr(13)
            Msg = Msg & " You can not edit this the record now" & Chr(13)
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
          Account_Code_dynamic = get_account_code_branch(127, my_branch)
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» ··«—«÷Ì    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
       
                End If
            End If
OptType(2).value = True

    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
    Rd(0).value = True
    Rd_Click (0)
  
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
  MySQL = "SELECT     dbo.TblBuyLanReEst.ID, dbo.TblBuyLanReEst.RecordDate, dbo.TblBuyLanReEst.Name, dbo.TblBuyLanReEst.NameE, dbo.TblBuyLanReEst.FullCode, "
  MySQL = MySQL & "                    dbo.TblBuyLanReEst.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBuyLanReEst.No_planned,"
  MySQL = MySQL & "                    dbo.TblBuyLanReEst.Area, dbo.TblBuyLanReEst.MeterPrice, dbo.TblBuyLanReEst.Total, dbo.TblBuyLanReEst.TitledeedNo, dbo.TblBuyLanReEst.PayType,"
  MySQL = MySQL & "                    dbo.TblBuyLanReEst.InstalNo, dbo.TblBuyLanReEst.FristDate, dbo.TblBuyLanReEst.Period, dbo.TblBuyLanReEst.PeriodType, dbo.TblBuyLanReEst.Remarks,"
  MySQL = MySQL & "                    dbo.TblBuyLanReEst.OwnerID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode,"
  MySQL = MySQL & "                    dbo.TblBuyLanReEstInsta.InstalDate, dbo.TblBuyLanReEstInsta.InstalNo AS InstalNoDet, dbo.TblBuyLanReEstInsta.ByLanID, dbo.TblBuyLanReEstInsta.Val,"
  MySQL = MySQL & "                    dbo.TblBuyLanReEstInsta.Remarks AS RemarksDet, dbo.TblBuyLanReEst.DesLocation, dbo.TblBuyLanReEst.Street, dbo.TblBuyLanReEst.DatePropertyDeed,"
  MySQL = MySQL & "                    dbo.TblBuyLanReEst.Unit, dbo.TblBuyLanReEst.Block, dbo.TblBuyLanReEst.PlateNo, dbo.TblBuyLanReEst.StreetNo, dbo.TblBuyLanReEst.NorthlengthStr,"
  MySQL = MySQL & "                    dbo.TblBuyLanReEst.SouthlengthStr, dbo.TblBuyLanReEst.EastlengthStr, dbo.TblBuyLanReEst.WestlengthStr, dbo.TblBuyLanReEst.Northlength,"
  MySQL = MySQL & "                    dbo.TblBuyLanReEst.Southlength, dbo.TblBuyLanReEst.Eastlength, dbo.TblBuyLanReEst.Westlength, dbo.TblBuyLanReEst.SchemeID,"
  MySQL = MySQL & "                    dbo.tblSchemes.name AS sCHname, dbo.tblSchemes.namee AS sCHnameE, dbo.TblBuyLanReEst.CountryID, dbo.TblCountriesData.CountryName,"
  MySQL = MySQL & "                    dbo.TblBuyLanReEst.CityID , dbo.TblCountriesGovernments.GovernmentName, dbo.TblBuyLanReEst.HyID, dbo.TblCountriesGovernmentsCities.CityName"
  MySQL = MySQL & "   FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCustemers RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBuyLanReEst INNER JOIN"
  MySQL = MySQL & "                    dbo.tblSchemes ON dbo.TblBuyLanReEst.SchemeID = dbo.tblSchemes.id LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCountriesGovernmentsCities ON dbo.TblBuyLanReEst.HyID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCountriesGovernments ON dbo.TblBuyLanReEst.CityID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCountriesData ON dbo.TblBuyLanReEst.CountryID = dbo.TblCountriesData.CountryID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBuyLanReEstInsta ON dbo.TblBuyLanReEst.ID = dbo.TblBuyLanReEstInsta.ByLanID ON dbo.TblCustemers.CusID = dbo.TblBuyLanReEst.OwnerID ON"
  MySQL = MySQL & "                    dbo.TblBranchesData.branch_id = dbo.TblBuyLanReEst.BranchId"
  MySQL = MySQL & "  Where (dbo.TblBuyLanReEst.id =" & val(TxtSerial1.Text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBuylandRealEstate.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBuylandRealEstateE.rpt"
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
    Wrap = Chr(13) + Chr(10)
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
    Me.Caption = "Buy Land and Real Estate  "
    Me.lbl(4).Caption = "Trans ID"
    Me.lbl(2).Caption = "Date"
    Me.lbl(7).Caption = "Branch"
    lbl(3).Caption = "Code"
    Me.Label1(2).Caption = Me.Caption
    lbl(1).Caption = "Des.Arabic"
    lbl(0).Caption = "Des.English"
    lbl(10).Caption = "No.Planned"
    lbl(12).Caption = "Area"
    lbl(13).Caption = "Meter Value"
    lbl(19).Caption = "Total Value"
    lbl(23).Caption = "Property Deed"
    Label1(11).Caption = "Payment"
    Label1(1).Caption = "Owner"
    Frame6.Caption = "Data of Payments"
    Label1(8).Caption = "No.Payments"
    Label1(9).Caption = "First Date"
    lbl(5).Caption = "Total"
    ISButton2.Caption = "Add"
   '''///////
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
   Label1(29).Caption = "Location"
   Label1(0).Caption = "Date"
   Label1(3).Caption = "Country"
   Label1(4).Caption = "City"
   Label1(6).Caption = "Street"
   Label1(5).Caption = "District"
   Label1(15).Caption = "Scheme"
   Label18.Caption = "Street No"
   Label17.Caption = "Plate No"
   Label16.Caption = "Block No."
   Label15.Caption = "No. Blvd."
   Frame10.Caption = ""
   ''//////
    lbl(21).Caption = "Remarks"

    
    lbl(11).Caption = "Payment Type"

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
lbl(6).Caption = 0
With Me.GridInstallments
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("PaymentNo"))) <> 0 Then
Counter = Counter + 1
.TextMatrix(i, .ColIndex("Ser")) = Counter
summation = summation + val(.TextMatrix(i, .ColIndex("PaymentValue")))
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
   My_SQL = "TblBuyLanReEst"
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

Private Sub TxtSouthLength_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSouthLength.Text, 0)
End Sub


Private Sub TxtTotalValue_Change()
lbl(17).Caption = WriteNo(val(Me.TxtTotalValue.Text), 0)
If Me.TxtModFlg.Text <> "R" Then
TxtTotalValue.Text = val(Me.TxtArea.Text) * val(TxtMeterValue.Text)
TxtTotalValue = Round(val(TxtTotalValue.Text), 2)
End If
End Sub

Private Sub txtWestlength_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.txtWestlength.Text, 0)
End Sub

Private Sub XPDtbTrans_Change()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.Text = ""
End If
End Sub
