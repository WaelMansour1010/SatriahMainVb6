VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FixedAssetsSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ «·«’Ê·"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   Icon            =   "FixedAssetSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   15270
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      ClipControls    =   0   'False
      Height          =   5175
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   -120
      Width           =   15270
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2E9E9&
         Height          =   1095
         Left            =   285
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   3480
         Width           =   14790
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4830
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   600
            Width           =   1185
         End
         Begin VB.TextBox TxtAssesetCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   12390
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   600
            Width           =   1185
         End
         Begin MSDataListLib.DataCombo DcbFromBranch 
            Bindings        =   "FixedAssetSearch.frx":030A
            Height          =   315
            Left            =   7680
            TabIndex        =   68
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
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
         Begin MSDataListLib.DataCombo DcbAssest 
            Bindings        =   "FixedAssetSearch.frx":031F
            Height          =   315
            Left            =   7680
            TabIndex        =   71
            Top             =   600
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
         Begin MSDataListLib.DataCombo DcbToBranch 
            Bindings        =   "FixedAssetSearch.frx":0334
            Height          =   315
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
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
            Bindings        =   "FixedAssetSearch.frx":0349
            Height          =   315
            Left            =   120
            TabIndex        =   75
            Top             =   600
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁ«∆„  »«·⁄„·Ì…"
            Height          =   285
            Index           =   11
            Left            =   6240
            TabIndex        =   77
            Top             =   600
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï ›—⁄"
            Height          =   285
            Index           =   10
            Left            =   6240
            TabIndex        =   76
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«’·"
            Height          =   285
            Index           =   8
            Left            =   13680
            TabIndex        =   72
            Top             =   600
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ ›—⁄"
            Height          =   285
            Index           =   7
            Left            =   13680
            TabIndex        =   69
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   2640
         Width           =   4695
         Begin MSDataListLib.DataCombo DcbBranch1 
            Bindings        =   "FixedAssetSearch.frx":035E
            Height          =   315
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
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
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   2
            Left            =   4080
            TabIndex        =   67
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame lbreg 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   2640
         Width           =   4575
         Begin MSComCtl2.DTPicker DtpDateFrom 
            Height          =   330
            Left            =   2280
            TabIndex        =   59
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   95420419
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DtpDateTo 
            Height          =   330
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   95420419
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï "
            Height          =   315
            Index           =   3
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ "
            Height          =   315
            Index           =   4
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Õ—ﬂ…"
            Height          =   195
            Index           =   1
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.Frame lbprocess 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   9960
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   2640
         Width           =   5115
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
            TabIndex        =   54
            Top             =   240
            Width           =   1155
         End
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
            TabIndex        =   53
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   315
            Index           =   5
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   315
            Index           =   6
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   240
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   195
            Index           =   14
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -360
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   2940
         Width           =   2175
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   2505
         Left            =   0
         TabIndex        =   51
         Top             =   120
         Width           =   15210
         _cx             =   26829
         _cy             =   4419
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
         Rows            =   50
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FixedAssetSearch.frx":0373
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   5175
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   -120
      Width           =   15285
      Begin VB.TextBox txtChaseeNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   12960
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   4440
         Width           =   1275
      End
      Begin VB.TextBox txtBoardNO 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   10800
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   4440
         Width           =   1275
      End
      Begin VB.ComboBox cStatus 
         Height          =   315
         ItemData        =   "FixedAssetSearch.frx":054E
         Left            =   120
         List            =   "FixedAssetSearch.frx":055E
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   2760
         Width           =   2085
      End
      Begin VB.ComboBox CBoDepreciation_Type_id 
         Height          =   315
         ItemData        =   "FixedAssetSearch.frx":059E
         Left            =   120
         List            =   "FixedAssetSearch.frx":05A8
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   3120
         Width           =   2085
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   12120
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   3540
         Width           =   2175
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·Â «Â·«ﬂ"
            Height          =   255
            Index           =   0
            Left            =   -480
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·Ì” ·Â"
            Height          =   315
            Index           =   1
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   0
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.TextBox TxtCustomerName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8700
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   3120
         Width           =   5535
      End
      Begin VB.TextBox XPTxtCusID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   12720
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2760
         Width           =   1515
      End
      Begin VB.TextBox TxtSalesFixed 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9780
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   4020
         Width           =   915
      End
      Begin VB.ComboBox DcbMVS 
         Height          =   315
         ItemData        =   "FixedAssetSearch.frx":05CB
         Left            =   9000
         List            =   "FixedAssetSearch.frx":05CD
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   4020
         Width           =   735
      End
      Begin VB.TextBox TxtAge 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6450
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   3660
         Width           =   405
      End
      Begin VB.TextBox TxtInstal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4050
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   3660
         Width           =   1155
      End
      Begin VB.ComboBox DcbMAhlak 
         Height          =   315
         ItemData        =   "FixedAssetSearch.frx":05CF
         Left            =   3240
         List            =   "FixedAssetSearch.frx":05D1
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox TxtValFixed 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4050
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   4020
         Width           =   1155
      End
      Begin VB.ComboBox DcbMVf 
         Height          =   315
         ItemData        =   "FixedAssetSearch.frx":05D3
         Left            =   3240
         List            =   "FixedAssetSearch.frx":05D5
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   4020
         Width           =   735
      End
      Begin VB.TextBox TxtBill 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6450
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   4020
         Width           =   1635
      End
      Begin VB.TextBox TxtDes 
         Alignment       =   1  'Right Justify
         Height          =   555
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   4380
         Width           =   5085
      End
      Begin VB.Frame Frame2 
         Caption         =   "ÃœÌœ/«›  «ÕÌ"
         Height          =   495
         Left            =   11880
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   3900
         Width           =   3255
         Begin VB.OptionButton OptNeworOpening 
            Alignment       =   1  'Right Justify
            Caption         =   "ÃœÌœ"
            Height          =   255
            Index           =   0
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton OptNeworOpening 
            Alignment       =   1  'Right Justify
            Caption         =   "«›  «ÕÌ"
            Height          =   255
            Index           =   1
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   120
            Width           =   975
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2505
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   15165
         _cx             =   26749
         _cy             =   4419
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
         Rows            =   50
         Cols            =   24
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FixedAssetSearch.frx":05D7
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
      Begin MSDataListLib.DataCombo DcEmployee 
         Height          =   315
         Left            =   3330
         TabIndex        =   21
         Top             =   3120
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCGroup 
         Height          =   315
         Left            =   3330
         TabIndex        =   22
         Top             =   2760
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DpPurchaseDate 
         Height          =   315
         Left            =   9000
         TabIndex        =   23
         Top             =   3660
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95420417
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker DTpEnd 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   3660
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95420417
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcCostCenter 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   4020
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcBranch 
         Height          =   315
         Left            =   8700
         TabIndex        =   40
         Top             =   2760
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   435
         Left            =   6450
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   4440
         Width           =   4245
         _cx             =   7488
         _cy             =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         BackColor       =   -2147483633
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
         Begin VB.TextBox txtNum4 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   0
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   0
            Width           =   555
         End
         Begin VB.TextBox txtLetter4 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2115
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   0
            Width           =   660
         End
         Begin VB.TextBox txtNum3 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   495
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   0
            Width           =   540
         End
         Begin VB.TextBox txtNum2 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   870
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   0
            Width           =   600
         End
         Begin VB.TextBox txtNum1 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1455
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   0
            Width           =   660
         End
         Begin VB.TextBox txtLetter3 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2625
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   0
            Width           =   585
         End
         Begin VB.TextBox txtLetter2 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3120
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   0
            Width           =   435
         End
         Begin VB.TextBox txtLetter1 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3525
            MaxLength       =   1
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   0
            Width           =   525
         End
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·‘«”Ì…"
         Height          =   255
         Index           =   15
         Left            =   13800
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   4440
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «··ÊÕ…"
         Height          =   255
         Index           =   12
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   4440
         Width           =   1395
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·«’·"
         Height          =   315
         Index           =   8
         Left            =   13920
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ«·… «·«’·"
         Height          =   255
         Index           =   118
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   2700
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·«Â·«ﬂ"
         Height          =   255
         Index           =   105
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·«’·"
         Height          =   315
         Index           =   1
         Left            =   13920
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   315
         Index           =   117
         Left            =   11160
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   2700
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»⁄ÂœÂ"
         Height          =   315
         Index           =   104
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   3120
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Ã„Ê⁄Â"
         Height          =   315
         Index           =   103
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… ‘—«¡ "
         Height          =   315
         Index           =   3
         Left            =   10920
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·‘—«¡"
         Height          =   255
         Index           =   128
         Left            =   10560
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   3780
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„— «·«› —«÷Ì ··«’· »«·‘Â—"
         Height          =   255
         Index           =   9
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   3660
         Width           =   2115
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… ﬁ”ÿ «·«Â·«ﬂ"
         Height          =   315
         Index           =   4
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ »œ«Ì… «·«Â·«ﬂ"
         Height          =   255
         Index           =   0
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   3660
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„—ﬂ“ «· ﬂ·›…"
         Height          =   195
         Index           =   13
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   4020
         Width           =   1125
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·«’· Œ—œÂ"
         Height          =   315
         Index           =   5
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·›« Ê—…"
         Height          =   315
         Index           =   6
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Ê’› «·«’·"
         Height          =   315
         Index           =   7
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   4440
         Width           =   975
      End
   End
   Begin VB.CheckBox XPChkSearchType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·«’· »«·ﬂ«„· "
      Height          =   435
      Left            =   12780
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5640
      Width           =   1575
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   2100
      TabIndex        =   1
      Top             =   5160
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
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
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   1110
      TabIndex        =   2
      Top             =   5160
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   180
      TabIndex        =   3
      Top             =   5160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰ ÌÃ… «·»ÕÀ"
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   2
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·«’·"
      Height          =   315
      Index           =   0
      Left            =   15840
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "FixedAssetsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim m_SearchType As Integer
Private m_DcboCustomers As DataCombo
Private m_RetrunType As Integer
Public branch_no As Integer

Public Sub GetData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = "SELECT     dbo.TblTransferAssetsDet.ID, dbo.TblTransferAssetsDet.TrnsAsseID, dbo.TblTransferAssetsDet.Remarks, dbo.TblTransferAssetsDet.PurchasePrice, "
    sql = sql & "                   dbo.TblTransferAssetsDet.AccDepreciation, dbo.TblTransferAssetsDet.FixedID, dbo.FixedAssets.Name, dbo.FixedAssets.Fullcode, dbo.FixedAssets.namee,"
    sql = sql & "                  dbo.TblTransferAssets.ID AS MainID, dbo.TblTransferAssets.Remarks AS MainRemarks, dbo.TblTransferAssets.RecordDate, dbo.TblTransferAssets.AssestID,"
    sql = sql & "                  FixedAssets_1.Name AS MainName, FixedAssets_1.namee AS MainNameE, FixedAssets_1.Fullcode AS MainFullcode, dbo.TblTransferAssets.BranchID,"
    sql = sql & "                  dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblTransferAssets.FromBranchID,"
    sql = sql & "                  TblBranchesData_1.branch_name AS Frombranch_name, TblBranchesData_1.branch_namee AS Frombranch_nameE, dbo.TblTransferAssets.ToBranchID,"
    sql = sql & "                  TblBranchesData_2.branch_name AS Tobranch_name, TblBranchesData_2.branch_namee AS Tobranch_nameE, dbo.TblTransferAssets.EmpID,"
    sql = sql & "                  dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1,"
    sql = sql & "                  dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
    sql = sql & "                  dbo.TblEmployee.Emp_Name3 , dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality"
    sql = sql & "  FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    sql = sql & "                  dbo.TblTransferAssets ON dbo.TblEmployee.Emp_ID = dbo.TblTransferAssets.EmpID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData TblBranchesData_2 ON dbo.TblTransferAssets.ToBranchID = TblBranchesData_2.branch_id LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData TblBranchesData_1 ON dbo.TblTransferAssets.FromBranchID = TblBranchesData_1.branch_id LEFT OUTER JOIN"
    sql = sql & "                  dbo.FixedAssets FixedAssets_1 ON dbo.TblTransferAssets.AssestID = FixedAssets_1.id LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.TblTransferAssets.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblTransferAssetsDet ON dbo.TblTransferAssets.ID = dbo.TblTransferAssetsDet.TrnsAsseID LEFT OUTER JOIN"
    sql = sql & "                  dbo.FixedAssets ON dbo.TblTransferAssetsDet.FixedID = dbo.FixedAssets.id"

    
       BolBegine = False
       StrWhere = ""
 

    
    
    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblTransferAssets.ID >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTransferAssets.ID >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblTransferAssets.ID <=" & val(Me.TxtIDTO.Text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblTransferAssets.ID <=" & val(Me.TxtIDTO.Text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblTransferAssets.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblTransferAssets.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblTransferAssets.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblTransferAssets.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcboEmpName.Text <> "" And (val(DcboEmpName.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblTransferAssets.EmpID =" & Me.DcboEmpName.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblTransferAssets.EmpID =" & Me.DcboEmpName.BoundText & ""
       End If
     End If
        If Me.DcbBranch1.Text <> "" And (val(DcbBranch1.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblTransferAssets.BranchID =" & Me.DcbBranch1.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblTransferAssets.BranchID =" & Me.DcbBranch1.BoundText & ""
       End If
     End If
          If Me.DcbFromBranch.Text <> "" And (val(DcbFromBranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblTransferAssets.FromBranchID =" & Me.DcbFromBranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblTransferAssets.FromBranchID =" & Me.DcbFromBranch.BoundText & ""
       End If
     End If
           If Me.DcbToBranch.Text <> "" And (val(DcbToBranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblTransferAssets.ToBranchID =" & Me.DcbToBranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblTransferAssets.ToBranchID =" & Me.DcbToBranch.BoundText & ""
       End If
     End If
               If Me.DcbAssest.Text <> "" And (val(DcbAssest.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblTransferAssetsDet.FixedID =" & Me.DcbAssest.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblTransferAssetsDet.FixedID =" & Me.DcbAssest.BoundText & ""
       End If
     End If

  '''''''''''''''''''''//////////////
  
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblTransferAssets.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
       If SystemOptions.UserInterface = ArabicInterface Then
                Me.XPLbl(2).Caption = "‰ ÌÃ… «·»ÕÀ : " & rs.RecordCount
            Else
                Me.XPLbl(2).Caption = "Search Results: " & rs.RecordCount
            End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid1
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
             If SystemOptions.UserInterface = ArabicInterface Then
                Me.XPLbl(2).Caption = "‰ ÌÃ… «·»ÕÀ : " & rs.RecordCount
            Else
                Me.XPLbl(2).Caption = "Search Results: " & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("NumIndex")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("MainID").value), "", rs("MainID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                  .TextMatrix(i, .ColIndex("purchaseprice")) = IIf(IsNull(rs("purchaseprice").value), 0, rs("purchaseprice").value)
                 .TextMatrix(i, .ColIndex("AccDepreciation")) = IIf(IsNull(rs("AccDepreciation").value), 0, rs("AccDepreciation").value)
                 .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
                
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("Frombranch_name")) = IIf(IsNull(rs("Frombranch_name").value), "", rs("Frombranch_name").value)
                .TextMatrix(i, .ColIndex("Tobranch_name")) = IIf(IsNull(rs("Tobranch_name").value), "", rs("Tobranch_name").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                Else
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                .TextMatrix(i, .ColIndex("Frombranch_name")) = IIf(IsNull(rs("Frombranch_nameE").value), "", rs("Frombranch_nameE").value)
                .TextMatrix(i, .ColIndex("Tobranch_name")) = IIf(IsNull(rs("Tobranch_nameE").value), "", rs("Tobranch_nameE").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Private Sub Cmd_Click(Index As Integer)
    Dim Msg As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
If RetrunType = 10 Then
GetData
Else
            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.Rows = 2
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.XPLbl(2).Caption = "‰ ÌÃ… «·»ÕÀ : " & rs.RecordCount
            Else
                Me.XPLbl(2).Caption = "Search Results: " & rs.RecordCount
            End If

            Retrive
            FG.SetFocus
         End If

        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            Opt(0).value = False
Opt(1).value = False
DpPurchaseDate.value = ""
DTpEnd.value = ""
DtpDateTo.value = ""
DtpDateFrom.value = ""

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub

Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.Text = ""
If Len(txtNum4.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
Cal_Board

End Sub
Private Sub Cal_Board()
    TxtBoardNO.Text = txtLetter1.Text & " " & txtLetter2.Text & " " & txtLetter3.Text & " " & txtLetter4.Text & " " & txtNum1.Text & " " & txtNum2.Text & " " & txtNum3.Text & " " & txtNum4.Text
End Sub
Private Sub txtNum4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board

End Sub
Private Sub txtNum3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtNum1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.Text = ""
If Len(txtNum1.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum2.SetFocus
End If
Cal_Board
End Sub
Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.Text = ""
If Len(txtNum2.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum3.SetFocus
End If
Cal_Board
End Sub
Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.Text = ""
If Len(txtNum3.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum4.SetFocus
End If
Cal_Board
End Sub
Private Sub Fg_Click()
    On Error GoTo ErrTrap

    If Not FG.TextMatrix(FG.Row, 1) = "" Then
        If Me.RetrunType = 0 Then
    
            FixedAssets.Retrive val(FG.TextMatrix(FG.Row, 0))
        ElseIf Me.RetrunType = 1 Then
'            FrmExpenses40.DcFixedAssets.BoundText = val(Fg.TextMatrix(Fg.Row, 0))
            
            ElseIf Me.RetrunType = 2 Then
            frmequipment.DcFixedAssets.BoundText = val(FG.TextMatrix(FG.Row, 0))
            
          ElseIf Me.RetrunType = 12 Then
            FrmExpenses40E.DcFixedAssets.BoundText = val(FG.TextMatrix(FG.Row, 0))
              
            ElseIf Me.RetrunType = 13 Then
            frmFixedAsseteports.DcFixedAssets.BoundText = val(FG.TextMatrix(FG.Row, 0))
               
            
           ElseIf Me.RetrunType = 3 Then
           FrmCars.DcFixedAssets.BoundText = val(FG.TextMatrix(FG.Row, 0))
            ElseIf Me.RetrunType = 4 Then
           frmequipment1.DcFixedAssets.BoundText = val(FG.TextMatrix(FG.Row, 0))
             
               ElseIf Me.RetrunType = 5 Then
            'frmequipment1.DcFixedAssets.BoundText = val(Fg.TextMatrix(Fg.Row, 0))
             
                       With FrmExpenses4.VSFlexGrid2
                ' FrmExpenses3.Fg_Journal_StartEdit .Row, 4, False
                .TextMatrix(.Row, .ColIndex("AccountName")) = FG.TextMatrix(FG.Row, 5)
                .TextMatrix(.Row, .ColIndex("id")) = FG.TextMatrix(FG.Row, 0)
        .TextMatrix(.Row, .ColIndex("AssetCode")) = FG.TextMatrix(FG.Row, FG.ColIndex("MemCode"))
       .TextMatrix(.Row, .ColIndex("GroupID")) = FG.TextMatrix(FG.Row, FG.ColIndex("GroupID"))
        .TextMatrix(.Row, .ColIndex("branch_id")) = FG.TextMatrix(FG.Row, FG.ColIndex("branch_id"))
        .TextMatrix(.Row, .ColIndex("AccountCode")) = get_FixedAsset_Account(val(.TextMatrix(.Row, .ColIndex("GroupID"))), val(.TextMatrix(.Row, .ColIndex("branch_id"))))
               
               
      
                   
           '     .TextMatrix(.Row, .ColIndex("des")) = Fg.TextMatrix(Fg.Row, 4)
             '  FrmExpenses4.VSFlexGrid2_AfterEdit .Row, 8
            End With
            
                        ElseIf Me.RetrunType = 6 Then
          FrmAccountingReport.DcFixedAssets.BoundText = val(FG.TextMatrix(FG.Row, 0))
                   
                        ElseIf Me.RetrunType = 7 Then
          FrmTransferAssets.DcbAssest.BoundText = val(FG.TextMatrix(FG.Row, 0))
                     ElseIf Me.RetrunType = 17 Then

          frmdriveassestMove.dcmboassest.BoundText = val(FG.TextMatrix(FG.Row, 0))
          
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("BoardNo")) = IIf(IsNull(rs("BoardNo").value), "", (rs("BoardNo").value))
                .TextMatrix(Num, .ColIndex("MemCode")) = IIf(IsNull(rs("Fullcode").value), "", (rs("Fullcode").value))
                .TextMatrix(Num, .ColIndex("MemNme")) = IIf(IsNull(rs("Name").value), "", Trim(rs("Name").value))
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", Trim(rs("id").value))
 
                .TextMatrix(Num, .ColIndex("GroupID")) = IIf(IsNull(rs("group_id").value), "", Trim(rs("group_id").value))
                
                .TextMatrix(Num, .ColIndex("branch_id")) = IIf(IsNull(rs("Branch_NO").value), "", Trim(rs("Branch_NO").value))
                 .TextMatrix(Num, .ColIndex("purchaseprice")) = IIf(IsNull(rs("purchaseprice").value), "", Trim(rs("purchaseprice").value))
                  .TextMatrix(Num, .ColIndex("PurchaseDate")) = IIf(IsNull(rs("PurchaseDate").value), "", Trim(rs("PurchaseDate").value))
                '''///
                
                .TextMatrix(Num, .ColIndex("KhordaPrice")) = IIf(IsNull(rs("KhordaPrice").value), "", Trim(rs("KhordaPrice").value))
                .TextMatrix(Num, .ColIndex("Notes")) = IIf(IsNull(rs("Notes").value), "", Trim(rs("Notes").value))
                .TextMatrix(Num, .ColIndex("PurchaseBillId")) = IIf(IsNull(rs("PurchaseBillId").value), "", Trim(rs("PurchaseBillId").value))
                
                .TextMatrix(Num, .ColIndex("Installmentvalue")) = IIf(IsNull(rs("Installmentvalue").value), "", Trim(rs("Installmentvalue").value))
                .TextMatrix(Num, .ColIndex("StartDepreciationDate")) = IIf(IsNull(rs("StartDepreciationDate").value), "", Trim(rs("StartDepreciationDate").value))
                .TextMatrix(Num, .ColIndex("DefaultAge")) = IIf(IsNull(rs("DefaultAge").value), "", Trim(rs("DefaultAge").value))
                If rs("HaveDepreciation").value = True Then
                .TextMatrix(Num, .ColIndex("HaveDepreciation")) = "‰⁄„"
                Else
                .TextMatrix(Num, .ColIndex("HaveDepreciation")) = "·«"
                End If
                
                         If rs("New_or_opening").value = 0 Then
                .TextMatrix(Num, .ColIndex("New_or_opening")) = "ÃœÌœ"
                Else
                .TextMatrix(Num, .ColIndex("New_or_opening")) = "«›  «ÕÌ"
                End If
                
                
                
                   If rs("Depreciation_Type_id").value = 1 Then
                .TextMatrix(Num, .ColIndex("Depreciation_Type_id")) = " «·ﬁ”ÿ «·À«Ì "
                Else
                .TextMatrix(Num, .ColIndex("Depreciation_Type_id")) = "  «·ﬁ”ÿ «·„ ‰«ﬁ’"
                End If
           
                  If rs("Status_id").value = 0 Then
                .TextMatrix(Num, .ColIndex("Status_id")) = " Ã«—Ì «·«Â·«ﬂ"
                ElseIf rs("Status_id").value = 1 Then
                .TextMatrix(Num, .ColIndex("Status_id")) = "  „ Êﬁ›"
                ElseIf rs("Status_id").value = 2 Then
                .TextMatrix(Num, .ColIndex("Status_id")) = " „ «· Œ·’ »«·»Ì⁄"
                ElseIf rs("Status_id").value = 3 Then
                .TextMatrix(Num, .ColIndex("Status_id")) = " „ «·«Â·«ﬂ »«· Œ—Ìœ"
                End If
           
                  If SystemOptions.UserInterface = ArabicInterface Then
                 ' .TextMatrix(Num, .ColIndex("account_name")) = IIf(IsNull(rs("account_name").value), "", Trim(rs("account_name").value))
                   .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", Trim(rs("branch_name").value))
                  .TextMatrix(Num, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", Trim(rs("GroupName").value))
                  .TextMatrix(Num, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", Trim(rs("Emp_Name").value))
                  Else
                 ' .TextMatrix(Num, .ColIndex("account_name")) = IIf(IsNull(rs("EnglishName").value), "", Trim(rs("EnglishName").value))
                  .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", Trim(rs("branch_namee").value))
                  .TextMatrix(Num, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupNamee").value), "", Trim(rs("GroupNamee").value))
                  .TextMatrix(Num, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", Trim(rs("Emp_Namee").value))
   
               End If
              
                ''///
                
            End With

            rs.MoveNext
        Next Num

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    Fg_Click
    Unload Me
End Sub
Private Sub txtLetter1_KeyPress(KeyAscii As Integer)
txtLetter1.Text = ""
If Len(txtLetter1.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        txtLetter2.SetFocus
End Select

End Sub
Private Sub txtLetter1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub Form_Activate()

  '  If Me.SearchType = 1 Then
  '      Me.Caption = "«·»ÕÀ ⁄‰ «·⁄„·«¡ Ê«·„Ê—œÌ‰"
  '      Me.XPLbl(1).Caption = "«·ﬂÊœ"
  '      Me.XPLbl(0).Caption = "«·√”„"
  '      XPChkSearchType.Caption = "«”„ «·‘Œ’ »«·ﬂ«„·"
'
'        With Me.Fg
'            .TextMatrix(0, .ColIndex("MemCode")) = "ﬂÊœ «·„Ê—œ «Ê «·⁄„Ì·"
'            .TextMatrix(0, .ColIndex("MemNme")) = "«”„ «·„Ê—œ «Ê «·⁄„Ì·"
'        End With
'
'    ElseIf Me.SearchType = 3 Then
'        Me.Caption = "«·»ÕÀ ⁄‰ »Ì«‰«  „ﬁ«Ê·Ì «·»«ÿ‰"
'        Me.XPLbl(1).Caption = "«·ﬂÊœ"
'        Me.XPLbl(0).Caption = "«·√”„"
'        XPChkSearchType.Caption = "«”„ «·‘Œ’ »«·ﬂ«„·"
'
'        With Me.Fg
'            .TextMatrix(0, .ColIndex("MemCode")) = "ﬂÊœ «·„ ⁄«„·"
'            .TextMatrix(0, .ColIndex("MemNme")) = "«”„ «·„ ⁄«„·"
'        End With
'
'    End If

End Sub
Private Sub DcbFromBranch_Click(Area As Integer)
loadAssest
End Sub
Sub loadAssest()
Dim StrSQL As String
   If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " Select  id,Name  From FixedAssets where HaveDepreciation=1   and PurchasePrice>0 "
        Else
        StrSQL = " Select  id,Namee  From FixedAssets where HaveDepreciation=1   and PurchasePrice>0 "
        End If
        fill_combo DcbAssest, StrSQL
End Sub
Sub GetAsseteCode_ID(Optional ByRef ID As Double = 0, Optional ByRef Fullcode As String = "", Optional Typ As Integer = 0)
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If Typ = 0 Then
sql = "select Fullcode  from FixedAssets where id=" & ID & " "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
Fullcode = IIf(IsNull(Rs7("Fullcode").value), "", Rs7("Fullcode").value)
Else
Fullcode = ""
End If
Else
sql = "select ID  from FixedAssets where Fullcode='" & Fullcode & "' "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ID = IIf(IsNull(Rs7("ID").value), 0, Rs7("ID").value)
Else
ID = 0
End If
End If
End Sub

Private Sub TxtAssesetCode_KeyPress(KeyAscii As Integer)
Dim AsseID As Double
If TxtAssesetCode.Text <> "" Then
If val(DcbFromBranch.BoundText) <> 0 Then
GetAsseteCode_ID AsseID, TxtAssesetCode.Text, 1
DcbAssest.BoundText = AsseID
End If
End If
End Sub

Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.Text = ""
If Len(txtLetter4.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.Text = ""
If Len(txtLetter3.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter4.SetFocus
End Select
Cal_Board
End Sub
Private Sub txtLetter2_KeyPress(KeyAscii As Integer)
txtLetter2.Text = ""
If Len(txtLetter2.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter3.SetFocus
End Select
Cal_Board
End Sub
Private Sub txtLetter4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
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
    TxtSearchCode.Text = EmpCode
End Sub
Private Sub DcbAssest_Change()
DcbAssest_Click (0)
End Sub

Private Sub DcbAssest_Click(Area As Integer)
Dim AsseCode1 As String
If val(DcbFromBranch.BoundText) <> 0 Then
If val(DcbAssest.BoundText) <> 0 Then
GetAsseteCode_ID val(DcbAssest.BoundText), AsseCode1, 0
TxtAssesetCode.Text = AsseCode1
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

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim StrSQL As String
        Dim Dcombos As New ClsDataCombos
Frame3.Visible = False
Frame4.Visible = False
If RetrunType = 10 Then
Me.Caption = "»ÕÀ ‰ﬁ· «·«’Ê·"
Frame4.Visible = True
Else
Me.Caption = "»ÕÀ «·«’Ê·"
Frame3.Visible = True
End If
loadAssest
Me.DcbMAhlak.AddItem ">"
Me.DcbMAhlak.AddItem "<"
Me.DcbMAhlak.AddItem ">="
Me.DcbMAhlak.AddItem "<="
Me.DcbMAhlak.AddItem "="
Me.DcbMVf.AddItem ">"
Me.DcbMVf.AddItem "<"
Me.DcbMVf.AddItem ">="
Me.DcbMVf.AddItem "<="
Me.DcbMVf.AddItem "="
Me.DcbMVS.AddItem ">"
Me.DcbMVS.AddItem "<"
Me.DcbMVS.AddItem ">="
Me.DcbMVS.AddItem "<="
Me.DcbMVS.AddItem "="

Opt(0).value = False
Opt(1).value = False
DpPurchaseDate.value = ""
DTpEnd.value = ""
DtpDateFrom.value = ""
DtpDateTo.value = ""

    Dcombos.GetFixedAssetsGroup DCGroup
Dcombos.GetBranches DcbBranch1
Dcombos.GetBranches DcbFromBranch
Dcombos.GetBranches DcbToBranch
Dcombos.GetEmployees Me.DcboEmpName
 StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL


    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
      Dcombos.GetBranches Dcbranch
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select Emp_ID,Emp_name  from TblEmployee order by Emp_name   "
 Else
 My_SQL = "  select Emp_ID,Emp_namee  from TblEmployee order by Emp_namee   "
 End If
    fill_combo dcEmployee, My_SQL
    Set rs = New ADODB.Recordset
    Dim BG As New ClsBackGroundPic
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    On Error GoTo ErrTrap

    If Me.SearchType = 0 Then
  '   StrSQL = " SELECT     dbo.FixedAssets.id, dbo.FixedAssets.code, dbo.FixedAssets.Name, dbo.FixedAssets.Branch_NO, dbo.TblBranchesData.branch_id, "
  '   StrSQL = StrSQL & "                 dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.FixedAssets.ReceiveDate, dbo.FixedAssets.CurrentValue, dbo.FixedAssets.Emp_id,"
  '   StrSQL = StrSQL & "                 dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee, dbo.FixedAssets.group_id,"
  '   StrSQL = StrSQL & "                 dbo.FixedAssetsGroup.GroupName, dbo.FixedAssets.Depreciation_Type_id, dbo.FixedAssets.LastDepreciationDate, dbo.FixedAssets.StartDepreciationDate,"
  '   StrSQL = StrSQL & "                 dbo.FixedAssets.NoOfInstallments, dbo.FixedAssets.EXEInstallments, dbo.FixedAssets.RemainInstallments, dbo.FixedAssets.PurchasePrice,"
  '   StrSQL = StrSQL & "                 dbo.FixedAssets.PurchaseDate, dbo.FixedAssets.PurchaseBillId, dbo.FixedAssets.KhordaPrice, dbo.FixedAssets.InstallmentValue, dbo.FixedAssets.New_or_opening,"
  '   StrSQL = StrSQL & "                 dbo.FixedAssets.Notes, dbo.FixedAssets.Fullcode, dbo.FixedAssets.prifix, dbo.FixedAssets.NoteSerial, dbo.FixedAssets.general_cost_center,"
  '   StrSQL = StrSQL & "                 dbo.markaas_taklefa.Code AS MarCode, dbo.markaas_taklefa.account_no, dbo.markaas_taklefa.account_name, dbo.FixedAssets.namee, dbo.FixedAssets.BiLLID,"
  '   StrSQL = StrSQL & "                 dbo.FixedAssets.MinusValue, dbo.FixedAssets.EndTest, dbo.FixedAssets.EndLicense, dbo.FixedAssets.Vendorid, dbo.FixedAssets.Contryid,"
  '   StrSQL = StrSQL & "                 dbo.FixedAssets.ChaseeNo, dbo.FixedAssets.Model, dbo.FixedAssets.SerialNo, dbo.FixedAssets.BoardNo, dbo.FixedAssets.opening_balance_voucher_id,"
  '   StrSQL = StrSQL & "                 dbo.FixedAssets.HaveDepreciation, dbo.FixedAssets.NoteID1, dbo.FixedAssets.NoteSerial1, dbo.FixedAssets.variance, dbo.FixedAssets.SalePrice,"
  '   StrSQL = StrSQL & "                 dbo.FixedAssets.NoteID, dbo.FixedAssets.Status_id, dbo.FixedAssets.DefaultAge, dbo.FixedAssets.AccDepreciation, dbo.FixedAssetsGroup.GroupNamee,"
  '   StrSQL = StrSQL & "                 dbo.markaas_taklefa.EnglishName"
  '  StrSQL = StrSQL & " FROM         dbo.FixedAssets LEFT OUTER JOIN"
  '  StrSQL = StrSQL & "                  dbo.markaas_taklefa ON dbo.FixedAssets.general_cost_center = dbo.markaas_taklefa.id LEFT OUTER JOIN"
  '  StrSQL = StrSQL & "                  dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID LEFT OUTER JOIN"
  '  StrSQL = StrSQL & "                  dbo.TblEmployee ON dbo.FixedAssets.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  '   StrSQL = StrSQL & "                 dbo.TblBranchesData ON dbo.FixedAssets.Branch_NO = dbo.TblBranchesData.branch_id"
  StrSQL = " SELECT     dbo.FixedAssets.id, dbo.FixedAssets.code, dbo.FixedAssets.Name, dbo.FixedAssets.Branch_NO, dbo.TblBranchesData.branch_id, "
  StrSQL = StrSQL & "                    dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.FixedAssets.ReceiveDate, dbo.FixedAssets.CurrentValue, dbo.FixedAssets.Emp_id,"
  StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee, dbo.FixedAssets.group_id,"
  StrSQL = StrSQL & "                     dbo.FixedAssetsGroup.GroupName, dbo.FixedAssets.Depreciation_Type_id, dbo.FixedAssets.LastDepreciationDate, dbo.FixedAssets.StartDepreciationDate,"
  StrSQL = StrSQL & "                     dbo.FixedAssets.NoOfInstallments, dbo.FixedAssets.EXEInstallments, dbo.FixedAssets.RemainInstallments, dbo.FixedAssets.PurchasePrice,"
  StrSQL = StrSQL & "                     dbo.FixedAssets.PurchaseDate, dbo.FixedAssets.PurchaseBillId, dbo.FixedAssets.KhordaPrice, dbo.FixedAssets.InstallmentValue, dbo.FixedAssets.New_or_opening,"
  StrSQL = StrSQL & "                     dbo.FixedAssets.Notes, dbo.FixedAssets.Fullcode, dbo.FixedAssets.prifix, dbo.FixedAssets.NoteSerial, dbo.FixedAssets.general_cost_center,"
  StrSQL = StrSQL & "                     dbo.FixedAssets.namee, dbo.FixedAssets.BiLLID, dbo.FixedAssets.MinusValue, dbo.FixedAssets.EndTest, dbo.FixedAssets.EndLicense, dbo.FixedAssets.Vendorid,"
  StrSQL = StrSQL & "                     dbo.FixedAssets.Contryid, dbo.FixedAssets.ChaseeNo, dbo.FixedAssets.Model, dbo.FixedAssets.SerialNo, dbo.FixedAssets.BoardNo,"
  StrSQL = StrSQL & "                     dbo.FixedAssets.opening_balance_voucher_id, dbo.FixedAssets.HaveDepreciation, dbo.FixedAssets.NoteID1, dbo.FixedAssets.NoteSerial1,"
  StrSQL = StrSQL & "                     dbo.FixedAssets.variance, dbo.FixedAssets.SalePrice, dbo.FixedAssets.NoteID, dbo.FixedAssets.Status_id, dbo.FixedAssets.DefaultAge,"
  StrSQL = StrSQL & "                     dbo.FixedAssets.AccDepreciation , dbo.FixedAssetsGroup.GroupNamee"
  StrSQL = StrSQL & " FROM         dbo.FixedAssets LEFT OUTER JOIN"
  StrSQL = StrSQL & "                      dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.FixedAssets.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.FixedAssets.Branch_NO = dbo.TblBranchesData.branch_id"
        Begin = False
    Else
        StrSQL = "select * From FixedAssets "
        Begin = False
    End If

 If Me.RetrunType = 5 Then
 StrSQL = StrSQL & " where New_or_opening=0 and PurchasePrice=0 order by Name"
   Begin = True
 End If
  If Me.RetrunType = 7 Then
 StrSQL = StrSQL & " where dbo.FixedAssets.Branch_NO=" & branch_no & "  order by Name"
   Begin = True
 End If
 
    If XPTxtCusID.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.Fullcode ='" & (XPTxtCusID.Text) & "'"
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.Fullcode ='" & (XPTxtCusID.Text) & "'"
            Begin = True
        End If
    End If

    If Trim(Me.txtCustomerName.Text) <> "" Then
        If XPChkSearchType.value = Checked Then
            If Begin = True Then
                StrWhere = StrWhere + " and dbo.FixedAssets.Name ='" & Trim(Me.txtCustomerName.Text) & "'"
            Else
                StrWhere = StrWhere + " where dbo.FixedAssets.Name ='" & Trim(Me.txtCustomerName.Text) & "'"
                Begin = True
            End If

        Else

            If Begin = True Then
                StrWhere = StrWhere + " and   dbo.FixedAssets.Name like '%" & Trim(txtCustomerName.Text) & "%'"
            Else
                StrWhere = StrWhere + " where dbo.FixedAssets.Name like '%" & Trim(txtCustomerName.Text) & "%'"
                Begin = True
            End If
        End If
    End If
 ''////
    If Me.SearchType = 0 Then
      If Not IsNull(DpPurchaseDate.value) Then
        If Begin = True Then
             StrWhere = StrWhere & " AND dbo.FixedAssets.PurchaseDate =" & SQLDate(Me.DpPurchaseDate.value, True) & ""
        Else
             StrWhere = StrWhere & " where dbo.FixedAssets.PurchaseDate =" & SQLDate(Me.DpPurchaseDate.value, True) & ""
            Begin = True
        End If
    End If
        If Opt(0).value = True Then
        If Begin = True Then
             StrWhere = StrWhere & " AND dbo.FixedAssets.HaveDepreciation =1"
        Else
             StrWhere = StrWhere & " where dbo.FixedAssets.HaveDepreciation =1"
            Begin = True
        End If
    End If
    
       If Opt(1).value = True Then
        If Begin = True Then
             StrWhere = StrWhere & " AND dbo.FixedAssets.HaveDepreciation =0"
        Else
             StrWhere = StrWhere & " where dbo.FixedAssets.HaveDepreciation =0"
            Begin = True
        End If
    End If
    
    
            If OptNeworOpening(0).value = True Then
        If Begin = True Then
             StrWhere = StrWhere & " AND dbo.FixedAssets.New_or_opening =0"
        Else
             StrWhere = StrWhere & " where dbo.FixedAssets.New_or_opening =0"
            Begin = True
        End If
    End If
    
       If OptNeworOpening(1).value = True Then
        If Begin = True Then
             StrWhere = StrWhere & " AND dbo.FixedAssets.New_or_opening =1"
        Else
             StrWhere = StrWhere & " where dbo.FixedAssets.New_or_opening =1"
            Begin = True
        End If
    End If
    
    
    
    
    
     If Not IsNull(DTpEnd.value) Then
        If Begin = True Then
             StrWhere = StrWhere & " AND dbo.FixedAssets.StartDepreciationDate =" & SQLDate(Me.DTpEnd.value, True) & ""
        Else
             StrWhere = StrWhere & " where dbo.FixedAssets.StartDepreciationDate =" & SQLDate(Me.DTpEnd.value, True) & ""
            Begin = True
        End If
    End If
        If txtChaseeNo.Text <> "" Then
        If Begin = True Then
        StrWhere = StrWhere & " AND REPLACE(dbo.FixedAssets.ChaseeNo, ' ', '')LIKE '%" & Replace(txtChaseeNo.Text, " ", "") & "%'"
        Else
           StrWhere = StrWhere & " where REPLACE(dbo.FixedAssets.ChaseeNo, ' ', '')LIKE '%" & Replace(txtChaseeNo.Text, " ", "") & "%'"
            Begin = True
        End If
    End If
    
      If TxtBoardNO.Text <> "" Then
        If Begin = True Then
        StrWhere = StrWhere & " AND REPLACE(dbo.FixedAssets.BoardNo, ' ', '')LIKE '%" & Replace(TxtBoardNO.Text, " ", "") & "%'"

           ' StrWhere = StrWhere + " and dbo.TblBranchesData.branch_id=" & val((dcBranch.BoundText)) & ""
        Else
               StrWhere = StrWhere & " where REPLACE(dbo.FixedAssets.BoardNo, ' ', '')LIKE '%" & Replace(TxtBoardNO.Text, " ", "") & "%'"

            'StrWhere = StrWhere + " where dbo.TblBranchesData.branch_id =" & val((dcBranch.BoundText)) & ""
            Begin = True
        End If
    End If
    
   If val(Dcbranch.BoundText) <> 0 And Dcbranch.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.TblBranchesData.branch_id=" & val((Dcbranch.BoundText)) & ""
        Else
            StrWhere = StrWhere + " where dbo.TblBranchesData.branch_id =" & val((Dcbranch.BoundText)) & ""
            Begin = True
        End If
    End If
       If val(DCGroup.BoundText) <> 0 And DCGroup.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.group_id=" & val((DCGroup.BoundText)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.group_id =" & val((DCGroup.BoundText)) & ""
            Begin = True
        End If
    End If
           If val(dcEmployee.BoundText) <> 0 And dcEmployee.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.Emp_id=" & val((dcEmployee.BoundText)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.Emp_id =" & val((dcEmployee.BoundText)) & ""
            Begin = True
        End If
    End If
             If val(DcCostCenter.BoundText) <> 0 And DcCostCenter.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.general_cost_center=" & val((DcCostCenter.BoundText)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.general_cost_center =" & val((DcCostCenter.BoundText)) & ""
            Begin = True
        End If
    End If
  If val(CBoDepreciation_Type_id.ListIndex) <> -1 And CBoDepreciation_Type_id.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.Depreciation_Type_id=" & val((CBoDepreciation_Type_id.ListIndex)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.Depreciation_Type_id =" & val((CBoDepreciation_Type_id.ListIndex)) & ""
            Begin = True
        End If
    End If
      If val(cStatus.ListIndex) <> -1 And cStatus.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.Status_id=" & val((cStatus.ListIndex)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.Status_id =" & val((cStatus.ListIndex)) & ""
            Begin = True
        End If
    End If
    
    If TxtAge.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.DefaultAge=" & val((TxtAge.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.DefaultAge =" & val((TxtAge.Text)) & ""
            Begin = True
        End If
    End If
        If TxtBill.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.PurchaseBillId='" & (TxtBill.Text) & "'"
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.PurchaseBillId ='" & (TxtBill.Text) & "'"
            Begin = True
        End If
    End If
            If TxtDes.Text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.Notes like '%" & Trim(TxtDes.Text) & "%'"
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.Notes like '%" & Trim(TxtDes.Text) & "%'"
            Begin = True
        End If
    End If
    
     If TxtSalesFixed.Text <> "" Then
     Select Case Me.DcbMVS.ListIndex
     Case 1
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.PurchasePrice > " & val((TxtSalesFixed.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.PurchasePrice > " & val((TxtSalesFixed.Text)) & ""
            Begin = True
        End If
     Case 0
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.PurchasePrice < " & val((TxtSalesFixed.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.PurchasePrice <" & val((TxtSalesFixed.Text)) & ""
            Begin = True
        End If
     Case 3
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.PurchasePrice >=" & val((TxtSalesFixed.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.PurchasePrice >=" & val((TxtSalesFixed.Text)) & ""
            Begin = True
        End If
     Case 2
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.PurchasePrice <= " & val((TxtSalesFixed.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.PurchasePrice <= " & val((TxtSalesFixed.Text)) & ""
            Begin = True
        End If
     Case 4
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.PurchasePrice= " & val((TxtSalesFixed.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.PurchasePrice = " & val((TxtSalesFixed.Text)) & ""
            Begin = True
        End If
        End Select
    End If
       If TxtValFixed.Text <> "" Then
     Select Case Me.DcbMVf.ListIndex
     Case 1
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.KhordaPrice > " & val((TxtValFixed.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.KhordaPrice > " & val((TxtValFixed.Text)) & ""
            Begin = True
        End If
     Case 0
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.KhordaPrice < " & val((TxtValFixed.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.KhordaPrice <" & val((TxtValFixed.Text)) & ""
            Begin = True
        End If
     Case 3
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.KhordaPrice >=" & val((TxtValFixed.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.KhordaPrice >=" & val((TxtValFixed.Text)) & ""
            Begin = True
        End If
     Case 2
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.KhordaPrice <= " & val((TxtValFixed.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.KhordaPrice <= " & val((TxtValFixed.Text)) & ""
            Begin = True
        End If
     Case 4
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.KhordaPrice= " & val((TxtValFixed.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.KhordaPrice = " & val((TxtValFixed.Text)) & ""
            Begin = True
        End If
        End Select
    End If
        If TxtInstal.Text <> "" Then
     Select Case Me.DcbMAhlak.ListIndex
     Case 1
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.InstallmentValue > " & val((TxtInstal.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.InstallmentValue > " & val((TxtInstal.Text)) & ""
            Begin = True
        End If
     Case 0
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.InstallmentValue < " & val((TxtInstal.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.InstallmentValue <" & val((TxtInstal.Text)) & ""
            Begin = True
        End If
     Case 3
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.InstallmentValue >=" & val((TxtInstal.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.InstallmentValue >=" & val((TxtInstal.Text)) & ""
            Begin = True
        End If
     Case 2
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.InstallmentValue <= " & val((TxtInstal.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.InstallmentValue <= " & val((TxtInstal.Text)) & ""
            Begin = True
        End If
     Case 4
        If Begin = True Then
            StrWhere = StrWhere + " and dbo.FixedAssets.InstallmentValue= " & val((TxtInstal.Text)) & ""
        Else
            StrWhere = StrWhere + " where dbo.FixedAssets.InstallmentValue = " & val((TxtInstal.Text)) & ""
            Begin = True
        End If
        End Select
    End If
    
    End If
''////

If Me.RetrunType = 3 Then

   If Begin = True Then
            StrWhere = StrWhere + " and ISEQUP=1"
        Else
            StrWhere = StrWhere + " where ISEQUP=1"
            Begin = True
        End If
        
  
End If

    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Not FG.TextMatrix(FG.Row, 1) = "" Then
            Fg_Click
        Else
            Cmd_Click (0)
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
If RetrunType = 10 Then
    Me.Caption = "Transfer Assets Search..."
 Else
 Me.Caption = "Assets Search..."
 End If
 lbl(15).Caption = "Chassis No.."
    XPLbl(1).Caption = "Code"
    XPLbl(0).Caption = " Name"
    lbl(14).Caption = "Trans No"
    lbl(1).Caption = "Trans Date"
    lbl(5).Caption = "From"
    lbl(4).Caption = "From"
    lbl(6).Caption = "To"
    lbl(3).Caption = "To"
    lbl(2).Caption = "Branch"
    lbl(7).Caption = "From Branch"
    lbl(10).Caption = "To Branch"
    lbl(8).Caption = "Assets "
    lbl(11).Caption = "Employee"
    
    lbl(117).Caption = "Branch"
    lbl(103).Caption = "Group"
     lbl(118).Caption = "Status"
     lbl(105).Caption = "Deprec. Type"
     Opt(0).RightToLeft = False
     lbl(104).Caption = "Employee"
     Opt(1).RightToLeft = False
     Opt(0).Caption = "Have Deprec."
     Opt(1).Caption = "Not Have"
     lbl(13).Caption = "Cost Center"
    XPLbl(2).Caption = "Search Results."
    XPLbl(4).Caption = "Insta. Value"
    XPLbl(5).Caption = "Value as Scrap "
    XPLbl(6).Caption = "Invoice .No"
    lbl(0).Caption = "Start Deprec"
    XPChkSearchType.Caption = "Math Complete Name"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"
lbl(9).Caption = "Default Age Month"
XPLbl(3).Caption = "Purc. Price"
lbl(128).Caption = "Purc. Date"
XPLbl(7).Caption = "Description"
XPLbl(8).Caption = "Name"
Frame2.Caption = "New/Opening"
OptNeworOpening(0).Caption = "New"
OptNeworOpening(1).Caption = "Opening"
OptNeworOpening(0).RightToLeft = False
OptNeworOpening(1).RightToLeft = False
    With Me.FG
        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
        .TextMatrix(0, .ColIndex("MemCode")) = " Code"
        .TextMatrix(0, .ColIndex("MemNme")) = "Name"
        .TextMatrix(0, .ColIndex("New_or_opening")) = "New/Opening "
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name "
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name "
        .TextMatrix(0, .ColIndex("Emp_Name")) = " "
        .TextMatrix(0, .ColIndex("purchaseprice")) = "Purchase Price "
        .TextMatrix(0, .ColIndex("PurchaseDate")) = "Purchase Date "
        .TextMatrix(0, .ColIndex("HaveDepreciation")) = "Have Depreciation "
        .TextMatrix(0, .ColIndex("DefaultAge")) = "Default Age "
        .TextMatrix(0, .ColIndex("Installmentvalue")) = "Installment Value "
        .TextMatrix(0, .ColIndex("Depreciation_Type_id")) = "Depreciation Type "
        .TextMatrix(0, .ColIndex("StartDepreciationDate")) = "Start DepreciationDate"
        .TextMatrix(0, .ColIndex("Status_id")) = "Status"
        .TextMatrix(0, .ColIndex("PurchaseBillId")) = "Invoice .No "
        .TextMatrix(0, .ColIndex("account_name")) = "Cost Ceter "
        .TextMatrix(0, .ColIndex("KhordaPrice")) = "Value as Scrap "
        .TextMatrix(0, .ColIndex("Notes")) = "Descriptions "
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee "
        .TextMatrix(0, .ColIndex("BoardNo")) = "Board Number "
    End With
    lbl(12).Caption = "Board No."
    With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
        .TextMatrix(0, .ColIndex("ID")) = "Trans No"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Trans Date "
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name "
        .TextMatrix(0, .ColIndex("Frombranch_name")) = "From Branch Name "
        .TextMatrix(0, .ColIndex("Tobranch_name")) = "To Branch Name "
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name "
        .TextMatrix(0, .ColIndex("Name")) = "Assets Name "
        .TextMatrix(0, .ColIndex("purchaseprice")) = "Value  "
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks  "
        .TextMatrix(0, .ColIndex("AccDepreciation")) = "Acc. Deprec "
   End With
End Sub

Public Property Get SearchType() As Integer
    SearchType = m_SearchType
End Property

Public Property Let SearchType(ByVal vNewValue As Integer)
    m_SearchType = vNewValue
    'm_SearchType=0 Search For Customers only
    'm_SearchType=1 Search For All table

End Property

Public Property Get DcboCustomers() As DataCombo
    Set DcboCustomers = m_DcboCustomers
End Property

Public Property Set DcboCustomers(ByVal vNewValue As DataCombo)
    Set m_DcboCustomers = vNewValue
End Property

Public Property Get RetrunType() As Integer
    RetrunType = m_RetrunType
End Property

Public Property Let RetrunType(ByVal vNewValue As Integer)
    m_RetrunType = vNewValue
End Property

Private Sub VSFlexGrid1_Click()
With VSFlexGrid1
FrmTransferAssets.FindRec val(.TextMatrix(.Row, .ColIndex("ID")))
End With
End Sub
