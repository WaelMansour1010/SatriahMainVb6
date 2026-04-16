VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmTravelTransactions 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· »Ì«‰«  «·—Õ·« "
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16665
   HelpContextID   =   280
   Icon            =   "FrmTravelsTransactions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   16665
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox XPTxtValView 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   118
      Top             =   6360
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   3375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   720
      Width           =   16455
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   510
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Text            =   "Text1"
         Top             =   990
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox CboPaymentType 
         Height          =   315
         Left            =   11640
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   480
         Width           =   3195
      End
      Begin VB.Frame FraNote 
         BackColor       =   &H00E2E9E9&
         Height          =   2205
         Left            =   11640
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   840
         Width           =   4635
         Begin VB.TextBox TxtChequeNumber 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   1320
            Width           =   3405
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   240
            Width           =   825
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   600
            Width           =   825
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   960
            Width           =   825
         End
         Begin MSComCtl2.DTPicker DtpChequeDueDate 
            Height          =   315
            Left            =   30
            TabIndex        =   85
            Top             =   1740
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   556
            _Version        =   393216
            Format          =   95420417
            CurrentDate     =   39614
         End
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   30
            TabIndex        =   86
            Top             =   960
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   0
            TabIndex        =   87
            Top             =   600
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCVendor 
            Height          =   315
            Left            =   0
            TabIndex        =   88
            Top             =   240
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·Œ“Ì‰…"
            Height          =   285
            Index           =   16
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·»‰ﬂ"
            Height          =   285
            Index           =   17
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   990
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   285
            Index           =   18
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
            Height          =   285
            Index           =   19
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   285
            Index           =   22
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.TextBox txtto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   79
         Top             =   1080
         Width           =   2715
      End
      Begin VB.TextBox TxtSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   13680
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   150
         Width           =   1215
      End
      Begin VB.TextBox txt_general_des 
         Alignment       =   1  'Right Justify
         Height          =   765
         Left            =   11640
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   77
         Top             =   3240
         Width           =   3435
      End
      Begin VB.TextBox txt_ORDER_NO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   16560
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   1590
         Width           =   2655
      End
      Begin VB.ComboBox CboPaymentType1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmTravelsTransactions.frx":038A
         Left            =   16920
         List            =   "FrmTravelsTransactions.frx":038C
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   510
         Width           =   3375
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   -240
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   1590
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   1590
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text7 
         DataField       =   "id"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   960
         TabIndex        =   72
         Text            =   "Text2"
         Top             =   1110
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TXT_A_NoteID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Text            =   "Text8"
         Top             =   5070
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   11520
         TabIndex        =   97
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   95420417
         CurrentDate     =   38784
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   7
         Left            =   -210
         TabIndex        =   98
         Top             =   5070
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "«·⁄—÷ «·ÃœÊ·Ï"
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
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin MSDataListLib.DataCombo dcproject 
         Height          =   315
         Left            =   17400
         TabIndex        =   99
         Top             =   1140
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcCostCenter 
         Bindings        =   "FrmTravelsTransactions.frx":038E
         Height          =   315
         Left            =   17520
         TabIndex        =   100
         Top             =   780
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
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
         Bindings        =   "FrmTravelsTransactions.frx":03A3
         Height          =   315
         Left            =   7680
         TabIndex        =   114
         Top             =   120
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo DCAccounts 
         Height          =   315
         Left            =   16920
         TabIndex        =   117
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Õ”«»"
         Height          =   285
         Index           =   26
         Left            =   16680
         RightToLeft     =   -1  'True
         TabIndex        =   116
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   10320
         RightToLeft     =   -1  'True
         TabIndex        =   115
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·—Õ·…"
         Height          =   285
         Index           =   4
         Left            =   14880
         RightToLeft     =   -1  'True
         TabIndex        =   112
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„’—Ê›« "
         Height          =   285
         Index           =   3
         Left            =   16560
         RightToLeft     =   -1  'True
         TabIndex        =   111
         Top             =   1110
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   12720
         RightToLeft     =   -1  'True
         TabIndex        =   110
         Top             =   135
         Width           =   795
      End
      Begin VB.Image ImgNote 
         Height          =   240
         Left            =   -240
         Picture         =   "FrmTravelsTransactions.frx":03B8
         Top             =   750
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‘—Ê⁄"
         Height          =   255
         Index           =   14
         Left            =   17040
         RightToLeft     =   -1  'True
         TabIndex        =   109
         Top             =   1140
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
         Height          =   255
         Index           =   15
         Left            =   14640
         RightToLeft     =   -1  'True
         TabIndex        =   108
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ ›« Ê—… «·„Ê—œ"
         Height          =   285
         Index           =   0
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   107
         Top             =   1110
         Width           =   1395
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
         Height          =   255
         Left            =   21720
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   810
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·‘—Õ «·⁄«„"
         Height          =   285
         Index           =   20
         Left            =   14760
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   3150
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÿ·»Ì…"
         Height          =   285
         Index           =   21
         Left            =   16200
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   1590
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·›« Ê—…"
         Height          =   285
         Index           =   23
         Left            =   17400
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   510
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1695
         Left            =   -120
         Top             =   510
         Width           =   1935
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Index           =   24
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   102
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Â–Â «·›« Ê—… ·‘—«¡ «·«’Ê· «·À«» … Ê ﬁÊ„ » —’Ìœ ﬁÌ„… ‘—«¡ «·«’· ›Ì „·› «·«’Ê· «·Ì«"
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
         Height          =   1500
         Index           =   25
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.OptionButton OptSort 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   195
      Index           =   1
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   240
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkLastAccount 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   195
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   195
      Index           =   0
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2340
      Left            =   21000
      TabIndex        =   47
      Top             =   4440
      Visible         =   0   'False
      Width           =   10755
      _cx             =   18971
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
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmTravelsTransactions.frx":0942
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
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3915
         Left            =   2550
         RightToLeft     =   -1  'True
         ScaleHeight     =   3915
         ScaleWidth      =   9405
         TabIndex        =   52
         Top             =   810
         Visible         =   0   'False
         Width           =   9405
         Begin VB.CommandButton Command3 
            Caption         =   "Call des"
            Height          =   255
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Add des"
            Height          =   255
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   3600
            Width           =   1350
         End
         Begin VB.TextBox txtcodesub 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   3600
            Width           =   855
         End
         Begin VB.TextBox TxtDese 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   1485
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   53
            Top             =   2040
            Width           =   8955
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   3900
            Left            =   120
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   0
            Width           =   10905
            _cx             =   19235
            _cy             =   6879
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   20.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   16777215
            ForeColor       =   4210688
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
            CaptionPos      =   7
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   7
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
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   1605
               Left            =   0
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   58
               Top             =   480
               Visible         =   0   'False
               Width           =   8955
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000C&
               Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
               ForeColor       =   &H0000C8FF&
               Height          =   315
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   0
               Width           =   2445
            End
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   255
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Height          =   495
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   495
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   3480
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Õœœ —ﬁ„ «·ﬁÌœ «·„—«œ ‰”Œ…"
         Height          =   1215
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   3720
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton Command5 
            Caption         =   "‰”Œ"
            Height          =   255
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   255
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   240
            Width           =   1335
         End
      End
      Begin VDSCOMBOLibCtl.SmartCombo SmartCombo1 
         Height          =   315
         Left            =   240
         TabIndex        =   63
         ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
         Top             =   480
         Visible         =   0   'False
         Width           =   2475
         _cx             =   1973752078
         _cy             =   1973748268
         Alignment       =   0
         Appearance      =   3
         AutoSearch      =   0   'False
         BackColor       =   -2147483624
         BackgroundColor =   -2147483633
         BorderColor     =   0
         BorderVisible   =   -1  'True
         Caption         =   "SmartCombo1"
         CaptionAlignment=   4
         CaptionBackColor=   -2147483633
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionForeColor=   -2147483630
         CaptionHeight   =   15
         CaptionOnTop    =   0   'False
         CaptionMultiLine=   0
         Checkbox3D      =   0   'False
         CheckboxAlignment=   5
         CheckboxBackColor=   16777215
         CheckboxSize    =   13
         CheckboxValue   =   0
         BrowsePictureAlignment=   5
         BrowsePictureStretchH=   0
         BrowsePictureStretchV=   0
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
         ForeColor       =   0
         Gap             =   0
         HideSelection   =   -1  'True
         Locked          =   0   'False
         MaxLength       =   0
         MultiLine       =   0
         OnFocus         =   3
         PasswordChar    =   ""
         Picture         =   "FrmTravelsTransactions.frx":0C1E
         PictureAlignment=   5
         PictureBackColor=   -2147483624
         PictureStretchH =   0
         PictureStretchV =   0
         Redraw          =   -1  'True
         ScrollBar       =   0
         Style           =   0
         Text            =   ""
         UnderLine       =   0   'False
         Enabled0        =   -1  'True
         Position0       =   0
         Tip0            =   "Caption"
         Visible0        =   0   'False
         Width0          =   90
         Enabled1        =   -1  'True
         Position1       =   1
         Tip1            =   ""
         Visible1        =   -1  'True
         Width1          =   32
         Enabled2        =   -1  'True
         Position2       =   2
         Tip2            =   "Check Box (Space, Ctrl + Space)"
         Visible2        =   0   'False
         Width2          =   16
         Enabled3        =   -1  'True
         Position3       =   3
         Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
         Visible3        =   -1  'True
         Width3          =   113
         Enabled4        =   -1  'True
         Position4       =   4
         Tip4            =   "Left Spinner (Alt + Left)"
         Visible4        =   0   'False
         Width4          =   16
         Enabled5        =   -1  'True
         Position5       =   5
         Tip5            =   "Right Spinner (Alt + Right)"
         Visible5        =   0   'False
         Width5          =   16
         Enabled6        =   -1  'True
         Position6       =   6
         Tip6            =   "Up Spinner (Ctrl + Up)"
         Visible6        =   0   'False
         Width6          =   16
         Enabled7        =   -1  'True
         Position7       =   7
         Tip7            =   "Down Spinner (Ctrl + Down)"
         Visible7        =   0   'False
         Width7          =   16
         Enabled8        =   -1  'True
         Position8       =   8
         Tip8            =   "Browse (Alt + Enter)"
         Visible8        =   0   'False
         Width8          =   16
         Enabled9        =   -1  'True
         Position9       =   9
         Tip9            =   " (Alt + Down, F4)"
         Visible9        =   -1  'True
         Width9          =   16
         Enabled10       =   -1  'True
         Position10      =   10
         Tip10           =   "Right Arrow (Alt + >)"
         Visible10       =   0   'False
         Width10         =   16
      End
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   7410
      Width           =   1905
   End
   Begin VB.TextBox Txt_Numorder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
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
      Height          =   1035
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   9420
      Width           =   6465
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   30
         Top             =   270
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboCreditSide 
         Height          =   315
         Left            =   90
         TabIndex        =   32
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   12
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   570
         Width           =   1485
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·› —… :"
         Height          =   315
         Index           =   13
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   315
         Index           =   11
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› œ«∆‰"
         Height          =   285
         Index           =   10
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› „œÌ‰"
         Height          =   285
         Index           =   9
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4320
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   645
      Left            =   17520
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   6420
      Width           =   2145
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   16575
      _cx             =   29236
      _cy             =   1349
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   24
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
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "FrmTravelsTransactions.frx":11B8
      Caption         =   "  ”ÃÌ· »Ì«‰«  «·—Õ·«   "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   6
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   0
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1695
         TabIndex        =   6
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTravelsTransactions.frx":1E92
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
         Height          =   375
         Index           =   2
         Left            =   630
         TabIndex        =   7
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTravelsTransactions.frx":222C
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
         Height          =   375
         Index           =   1
         Left            =   2220
         TabIndex        =   8
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTravelsTransactions.frx":25C6
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
         Height          =   375
         Index           =   3
         Left            =   1155
         TabIndex        =   9
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTravelsTransactions.frx":2960
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin MSAdodcLib.Adodc numbering 
         Height          =   585
         Left            =   4680
         Top             =   0
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1032
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   " Õ—Ìﬂ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc detect_no 
         Height          =   585
         Left            =   2640
         Top             =   0
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1032
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   " Õ—Ìﬂ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label LblShortcutKeys 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ÃœÌœ F12 Or Enter ,  ⁄œÌ· F11 , Õ›Ÿ F10 ,  —«Ã⁄ F9 ,Õ–› F8 ,»ÕÀ F7 "
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
         Height          =   195
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   510
         Width           =   5445
      End
   End
   Begin MSDataListLib.DataCombo XPCboExpensesType 
      Height          =   315
      Left            =   16920
      TabIndex        =   1
      Top             =   2760
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   7680
      TabIndex        =   13
      Top             =   7410
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   7980
      TabIndex        =   19
      Top             =   6840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   1
      Left            =   7080
      TabIndex        =   20
      Top             =   6840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   2
      Left            =   6270
      TabIndex        =   21
      Top             =   6840
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   3
      Left            =   5115
      TabIndex        =   22
      Top             =   6870
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   4
      Left            =   4200
      TabIndex        =   23
      Top             =   6870
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   24
      Top             =   6870
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdHelp 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   1080
      TabIndex        =   25
      Top             =   6870
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   5
      Left            =   3150
      TabIndex        =   26
      Top             =   6870
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
      Height          =   2340
      Left            =   17880
      TabIndex        =   37
      Top             =   4440
      Visible         =   0   'False
      Width           =   10800
      _cx             =   19050
      _cy             =   4128
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
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmTravelsTransactions.frx":2CFA
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
         ScaleWidth      =   2925
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   2925
         Begin VB.TextBox TxtDes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   1125
            Left            =   30
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   41
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
            TabIndex        =   42
            Top             =   0
            Width           =   2445
         End
      End
      Begin VDSCOMBOLibCtl.SmartCombo CboDes 
         Height          =   315
         Left            =   240
         TabIndex        =   43
         ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
         Top             =   600
         Visible         =   0   'False
         Width           =   2955
         _cx             =   1973752924
         _cy             =   1973748268
         Alignment       =   0
         Appearance      =   3
         AutoSearch      =   0   'False
         BackColor       =   -2147483624
         BackgroundColor =   -2147483633
         BorderColor     =   0
         BorderVisible   =   -1  'True
         Caption         =   "SmartCombo1"
         CaptionAlignment=   4
         CaptionBackColor=   -2147483633
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionForeColor=   -2147483630
         CaptionHeight   =   15
         CaptionOnTop    =   0   'False
         CaptionMultiLine=   0
         Checkbox3D      =   0   'False
         CheckboxAlignment=   5
         CheckboxBackColor=   16777215
         CheckboxSize    =   13
         CheckboxValue   =   0
         BrowsePictureAlignment=   5
         BrowsePictureStretchH=   0
         BrowsePictureStretchV=   0
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
         ForeColor       =   0
         Gap             =   0
         HideSelection   =   -1  'True
         Locked          =   0   'False
         MaxLength       =   0
         MultiLine       =   0
         OnFocus         =   3
         PasswordChar    =   ""
         Picture         =   "FrmTravelsTransactions.frx":2E60
         PictureAlignment=   5
         PictureBackColor=   -2147483624
         PictureStretchH =   0
         PictureStretchV =   0
         Redraw          =   -1  'True
         ScrollBar       =   0
         Style           =   0
         Text            =   ""
         UnderLine       =   0   'False
         Enabled0        =   -1  'True
         Position0       =   0
         Tip0            =   "Caption"
         Visible0        =   0   'False
         Width0          =   90
         Enabled1        =   -1  'True
         Position1       =   1
         Tip1            =   ""
         Visible1        =   -1  'True
         Width1          =   32
         Enabled2        =   -1  'True
         Position2       =   2
         Tip2            =   "Check Box (Space, Ctrl + Space)"
         Visible2        =   0   'False
         Width2          =   16
         Enabled3        =   -1  'True
         Position3       =   3
         Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
         Visible3        =   -1  'True
         Width3          =   145
         Enabled4        =   -1  'True
         Position4       =   4
         Tip4            =   "Left Spinner (Alt + Left)"
         Visible4        =   0   'False
         Width4          =   16
         Enabled5        =   -1  'True
         Position5       =   5
         Tip5            =   "Right Spinner (Alt + Right)"
         Visible5        =   0   'False
         Width5          =   16
         Enabled6        =   -1  'True
         Position6       =   6
         Tip6            =   "Up Spinner (Ctrl + Up)"
         Visible6        =   0   'False
         Width6          =   16
         Enabled7        =   -1  'True
         Position7       =   7
         Tip7            =   "Down Spinner (Ctrl + Down)"
         Visible7        =   0   'False
         Width7          =   16
         Enabled8        =   -1  'True
         Position8       =   8
         Tip8            =   "Browse (Alt + Enter)"
         Visible8        =   0   'False
         Width8          =   16
         Enabled9        =   -1  'True
         Position9       =   9
         Tip9            =   " (Alt + Down, F4)"
         Visible9        =   -1  'True
         Width9          =   16
         Enabled10       =   -1  'True
         Position10      =   10
         Tip10           =   "Right Arrow (Alt + >)"
         Visible10       =   0   'False
         Width10         =   16
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   255
      Left            =   9120
      TabIndex        =   38
      Top             =   6960
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "„—«ﬂ“ «· ﬂ·›…"
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
      MICON           =   "FrmTravelsTransactions.frx":33FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   8
      Left            =   2160
      TabIndex        =   44
      Top             =   6960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   9
      Left            =   5640
      TabIndex        =   45
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·‘Ìﬂ"
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ALLButtonS.ALLButton CmdRemove 
      Height          =   375
      Left            =   9600
      TabIndex        =   46
      Tag             =   "Delete Row"
      Top             =   6360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ–› ”ÿ—"
      ENAB            =   0   'False
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
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   255
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTravelsTransactions.frx":3416
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   10
      Left            =   3480
      TabIndex        =   67
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
      Height          =   2340
      Left            =   120
      TabIndex        =   113
      Top             =   3960
      Width           =   10800
      _cx             =   19050
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
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   2
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmTravelsTransactions.frx":3432
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
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   390
      Index           =   8
      Left            =   9345
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   7425
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   255
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Height          =   405
      Left            =   3390
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6420
      Width           =   6015
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   7410
      Width           =   555
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   7410
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "/"
      Height          =   435
      Index           =   6
      Left            =   690
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   7410
      Width           =   165
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
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
      Height          =   435
      Index           =   7
      Left            =   1500
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   7410
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«Ã„«·Ì"
      Height          =   285
      Index           =   2
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   6480
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·«„—"
      Height          =   285
      Index           =   5
      Left            =   17520
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2520
      Width           =   1515
   End
End
Attribute VB_Name = "FrmTravelTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim numbering_type As Integer
Dim departement_name  As String
Dim branch_no  As String
Dim RsNotes As ADODB.Recordset
Dim BolEditOnMainAccounts As Boolean

 Function CuurentLogdata(Optional Currentmode As String)
LogTextA = "    ‘«‘… " & ScreenNameArabic _
                   & Chr(13) & "—ﬁ„ «·›« Ê—… " & TxtSerial1.text _
                   & Chr(13) & "   «· «—ÌŒ  " & XPDtbTrans _
                   & Chr(13) & "   «·›—⁄ " & dcBranch _
                   & Chr(13) & "   ÿ—Ìﬁ… «·œ›⁄  " & CboPayMentType _
                   & Chr(13) & "   «·Œ“Ì‰… " & DcboBox _
                   & Chr(13) & "   «·»‰ﬂ  " & DcboBankName _
                   & Chr(13) & "   —ﬁ„ «·‘Ìﬂ " & TxtChequeNumber _
                   & Chr(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate _
                   & Chr(13) & "   «·„Ê—œ  " & DCVendor _
                   & Chr(13) & " —ﬁ„ ›« Ê—… «·„Ê—œ" & TXTTo _
                   & Chr(13) & " «·Õ”«»  " & DCAccounts _
                   & Chr(13) & "   «·‘—Õ «·⁄«„  " & txt_general_des _
                    & Chr(13) & "   «Ã„«·Ì «·”‰œ    " & XPTxtValView _

                     
 LogTextE = "    Screen  " & ScreenNameEnglish _
                      & Chr(13) & " Bill . No " & TxtSerial1.text _
                   & Chr(13) & "   Date  " & XPDtbTrans _
                   & Chr(13) & "   Branch " & dcBranch _
                   & Chr(13) & "  Payment Type  " & CboPayMentType _
                   & Chr(13) & "   Box " & DcboBox _
                   & Chr(13) & "   Bank  " & DcboBankName _
                   & Chr(13) & "   Cheque No:   " & TxtChequeNumber _
                   & Chr(13) & "   Supplier  " & DCVendor _
                   & Chr(13) & "Supill No plier B" & TXTTo _
                   & Chr(13) & " Account  " & DCAccounts _
                   & Chr(13) & "  Remarks  " & txt_general_des _
                    & Chr(13) & "   Vchr Total   " & XPTxtValView _

                     
         If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 300, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, , , TxtSerial, TxtSerial1
    Else
    AddToLogFile CInt(user_id), 300, Date, Time, LogTextA, LogTextE, Me.name, "D", , , TxtSerial, TxtSerial1
    End If
    
End Function


Function saveChequeBoxContents1(NoteID As Double)
If SystemOptions.banks_Accounts3 = False Then Exit Function
 Dim i As Integer
 Dim rs As New ADODB.Recordset
 
 Dim StrSQL As String

StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords
 
 
rs.Open "TblChecqueBoxContent1", Cn, adOpenStatic, adLockOptimistic, adCmdTable

 If CboPayMentType.ListIndex = 1 Then
        rs.AddNew
        rs("noteid").value = NoteID
     
        rs("RecordDate").value = XPDtbTrans.value
        rs("DueDate").value = DtpChequeDueDate.value
        rs("BankID").value = Val(DcboBankName.BoundText)
        rs("BankName").value = DcboBankName.text
        
        rs("ChequeNo").value = TxtChequeNumber.text
        rs("ChequeValue").value = Val(XPTxtVal.text)
    
        rs("Remarks").value = Me.DcboDebitSide.text
        rs("Payed").value = 0
       
      rs("DepitAccount").value = (DcboDebitSide.BoundText)
      rs("notes_all").value = NoteID
      
        rs.update
  End If
rs.Close
End Function

Private Sub ALLButton1_Click()
On Error GoTo ErrTrap
If DcCostCenter.BoundText <> "" Then

MsgBox "·«Ì„ﬂ‰ «· Ê“Ì⁄ ⁄·Ï „—«ﬂ“ «· ﬂ·›… ·«‰ﬂ «Œ —   Ê“Ì⁄ ⁄«„ ⁄·Ï „—ﬂ“  ﬂ·›… „Õœœ", vbCritical
Exit Sub
End If

Dim opr_id As Double
If Not IsNumeric(Text1.text) Then Exit Sub
'If Me.TxtModFlg.text = "N" Then
opr_id = Val(Me.Text1.text)
'Else
'opr_id = TxtDEV_NO.text
'End If


       If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) = "" Then
             If Not Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE"))) = 0 Then


            marakes_taklefa_tawze3.Show
            
            marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
            marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
            marakes_taklefa_tawze3.LineNo = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        
                 Else
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… ", vbCritical
         Else
         MsgBox "Enter Value First ", vbCritical
         End If
         Exit Sub
            End If
            
            End If

            marakes_taklefa_tawze3.opr_type = "”‰œ ’—›"
            marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
            marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
            marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
            marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
            marakes_taklefa_tawze3.Adodc3.Refresh
        '    Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"


Exit Sub
ErrTrap:
End Sub

Private Sub CboDes_AfterAutoCloseUp()
PutData
CboDes.Visible = False
End Sub

Private Sub CboPayMentType_Change()
If Me.TxtModFlg.text = "E" Then
    DcboBankName.text = ""
    TxtChequeNumber.text = ""
    Me.DcboBox.text = ""
    DCVendor.text = ""
    DCAccounts.text = ""
End If


If Me.CboPayMentType.ListIndex = 0 Then
    Me.lbl(16).Enabled = True
    Me.DcboBox.Enabled = True
    Me.lbl(19).Enabled = False
    Me.lbl(18).Enabled = False
    Me.lbl(17).Enabled = False
    Me.DcboBankName.Enabled = False
    Me.TxtChequeNumber.Enabled = False
    Me.DtpChequeDueDate.Enabled = False
     Me.DCVendor.Enabled = False
         DcboBankName.text = ""
    TxtChequeNumber.text = ""
    
ElseIf Me.CboPayMentType.ListIndex = 1 Then
    Me.lbl(16).Enabled = False
    Me.DcboBox.Enabled = False
      DcboBox.text = ""
    Me.lbl(19).Enabled = True
    Me.lbl(18).Enabled = True
    Me.lbl(17).Enabled = True
    Me.DcboBankName.Enabled = True
    Me.TxtChequeNumber.Enabled = True
    Me.DtpChequeDueDate.Enabled = True
     Me.DCVendor.Enabled = False
        If SystemOptions.UserInterface = ArabicInterface Then
  lbl(18).Caption = "—ﬁ„ «·‘Ìﬂ "
    lbl(19).Caption = " «—ÌŒ «·«” Õﬁ«ﬁ"
    
   Else
    lbl(18).Caption = "Cheque No"
    lbl(19).Caption = "Due Date"
    End If
    
ElseIf Me.CboPayMentType.ListIndex = 2 Then
    Me.lbl(16).Enabled = True
    Me.DcboBox.Enabled = True
    Me.lbl(19).Enabled = False
    Me.lbl(18).Enabled = False
    Me.lbl(17).Enabled = False
    Me.DcboBankName.Enabled = False
    Me.TxtChequeNumber.Enabled = False
    Me.DtpChequeDueDate.Enabled = False
    Me.DcboBox.Enabled = False
    Me.DCVendor.Enabled = True
    
    ElseIf Me.CboPayMentType.ListIndex = 3 Then
    Me.lbl(16).Enabled = False
    Me.DcboBox.Enabled = False
      DcboBox.text = ""
    Me.lbl(19).Enabled = True
    Me.lbl(18).Enabled = True
    Me.lbl(17).Enabled = True
    Me.DcboBankName.Enabled = True
    Me.TxtChequeNumber.Enabled = True
    Me.DtpChequeDueDate.Enabled = True
     Me.DCVendor.Enabled = False
   If SystemOptions.UserInterface = ArabicInterface Then
  lbl(18).Caption = "—ﬁ„ «·ÕÊ«·… "
    lbl(19).Caption = " «—ÌŒÂ«"
       Else
    lbl(18).Caption = "Transfer No"
    lbl(19).Caption = "Date"
    End If
    
    ElseIf Me.CboPayMentType.ListIndex = 5 Then
    Me.lbl(16).Enabled = False
    Me.DcboBox.Enabled = False
      DcboBox.text = ""
    Me.lbl(19).Enabled = True
    Me.lbl(18).Enabled = True
    Me.lbl(17).Enabled = True
    Me.DcboBankName.Enabled = True
    Me.TxtChequeNumber.Enabled = True
    Me.DtpChequeDueDate.Enabled = True
     Me.DCVendor.Enabled = False
   If SystemOptions.UserInterface = ArabicInterface Then
  lbl(18).Caption = "—ﬁ„ «·‘Ìﬂ "
    lbl(19).Caption = " «—ÌŒÂ"
       
       
   Else
    lbl(18).Caption = "Cheque No"
    lbl(19).Caption = "Date"
    End If
    
    
    
ElseIf Me.CboPayMentType.ListIndex = 4 Then
    Me.lbl(16).Enabled = True
    Me.DcboBox.Enabled = False
    Me.lbl(19).Enabled = False
    Me.lbl(18).Enabled = False
    Me.lbl(17).Enabled = False
    Me.DcboBankName.Enabled = False
    Me.TxtChequeNumber.Enabled = False
    Me.DtpChequeDueDate.Enabled = False
     Me.DCVendor.Enabled = False
         DcboBankName.text = ""
    TxtChequeNumber.text = ""
    DCVendor.BoundText = ""
    DcboBox.BoundText = ""
    DcboBankName.BoundText = ""
    DCAccounts.Enabled = True
    
Else
    Me.lbl(16).Enabled = True
    Me.DcboBox.Enabled = True
    Me.lbl(19).Enabled = False
    Me.lbl(18).Enabled = False
    Me.lbl(17).Enabled = False
    Me.DcboBankName.Enabled = False
    Me.TxtChequeNumber.Enabled = False
    Me.DtpChequeDueDate.Enabled = False
      Me.DCVendor.Enabled = False
End If

End Sub

Private Sub CboPayMentType_Click()
CboPayMentType_Change
End Sub

Function setfoxy()
   Text1.text = CStr(new_id("foxy", "id", "", True))

   Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable

 
 rs("id").value = Text1.text
 
 rs.update
 
    
End Function

Private Sub CboPaymentType1_Change()

If Me.CboPaymentType1.ListIndex = 0 Then
Fg_Journal.Visible = True
VSFlexGrid1.Visible = False

ElseIf Me.CboPaymentType1.ListIndex = 1 Then
Fg_Journal.Visible = False
VSFlexGrid1.Visible = True
End If

End Sub

Private Sub CboPaymentType1_Click()
CboPaymentType1_Change
End Sub

Private Sub Cmd_Click(Index As Integer)
On Error GoTo ErrTrap
Select Case Index
    Case 0
        If DoPremis(Do_New, Me.name, True) = False Then
            Exit Sub
        End If
       
        TxtModFlg.text = "N"
        clear_all Me
        DcCostCenter.text = ""
         CboPaymentType1.ListIndex = 2
       ' XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
       ' Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
        
        Me.DCboUserName.BoundText = user_id
'        XPDtbTrans.SetFocus
Fg_Journal.Visible = False
VSFlexGrid1.Visible = False

        Fg_Journal.Clear flexClearScrollable, flexClearEverything
          Fg_Journal.Rows = 3
          Fg_Journal.Enabled = True
          
           VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
          VSFlexGrid1.Rows = 2
          VSFlexGrid1.Enabled = True
          
          
           VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
          VSFlexGrid2.Rows = 2
          VSFlexGrid2.Enabled = True
          
          DtpChequeDueDate.value = Date
          setfoxy
          Me.dcBranch.BoundText = branch_id
    Case 1
                  Dim Msg  As String
                    If SystemOptions.banks_Accounts3 = True Then
        If ChequeBoxOperations1(Val(Me.XPTxtID)) = False Then
            Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–… «·⁄„·Ì…"
            Msg = Msg & Chr(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    End If
    
        If DoPremis(Do_Edit, Me.name, True) = False Then
            Exit Sub
        End If
        TxtModFlg.text = "E"
        Me.DCboUserName.BoundText = user_id
        Fg_Journal.Rows = Fg_Journal.Rows + 1
        Fg_Journal.Enabled = True
        VSFlexGrid1.Rows = VSFlexGrid1.Rows + 1
        VSFlexGrid1.Enabled = True
       
               VSFlexGrid2.Rows = VSFlexGrid2.Rows + 1
        VSFlexGrid2.Enabled = True
        CuurentLogdata
    Case 2
      
        If Trim(dcBranch.BoundText) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify branch"
            Else
                Msg = "Õœœ «·›—⁄"
            End If
   MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    dcBranch.SetFocus
    SendKeys "{F4}"
    Screen.MousePointer = vbDefault
    Exit Sub
    End If
 my_branch = Me.dcBranch.BoundText


    
    
    
    DcboBox_Change
DcboBankName_Change
DCVendor_Change
    DCAccounts_Change
        SaveData
         
           
    Case 3
       Undo
    Case 4
        If DoPremis(Do_Delete, Me.name, True) = False Then
            Exit Sub
        End If
        Del_Trans
    Case 5
        If DoPremis(Do_Search, Me.name, True) = False Then
            Exit Sub
        End If
        Load FrmNotesSearch
        FrmNotesSearch.SearchType = 300
        FrmNotesSearch.Show vbModal
    Case 6
        Unload Me
    Case 7
        ViewDataList
    Case 8
         print_report (TxtSerial.text)
    Case 9
         print_Cheque TxtChequeNumber.text, get_Cheque_report_no(Val(DcboBankName.BoundText)), TxtSerial.text
    Case 10
    ShowGL_cc TxtSerial.text, , 200
    
End Select
Exit Sub
ErrTrap:
End Sub
Function print_Cheque(Optional ChqueNum As String = "", Optional report_no As String = "", Optional serial As String)
    hide_logo = True
Dim MySQL As String
Dim RsData As New ADODB.Recordset
Dim xApp As New CRAXDRT.Application
Dim xReport As CRAXDRT.Report
Dim CViewer As ClsReportViewer
Dim StrReportTitle As String
Dim StrFileName As String
Dim Msg As String

MySQL = "Select * From Expanses_Order  where ChqueNum='" & ChqueNum & "' and noteserial='" & TxtSerial & "'"

 

If SystemOptions.UserInterface = ArabicInterface Then
    StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
Else
    StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
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
 'MsgBox ToHijriDate(Date)

 xReport.ParameterFields(5).AddCurrentValue Mid(ToHijriDate(DtpChequeDueDate.value), 1, 2)
 xReport.ParameterFields(6).AddCurrentValue Mid(ToHijriDate(DtpChequeDueDate.value), 4, 2)
 xReport.ParameterFields(7).AddCurrentValue Mid(ToHijriDate(DtpChequeDueDate.value), 9, 2)
 

  xReport.ParameterFields(8).AddCurrentValue Mid(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 1, 2)
 xReport.ParameterFields(9).AddCurrentValue Mid(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 4, 2)
 xReport.ParameterFields(10).AddCurrentValue Mid(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 9, 2)
 xReport.ParameterFields(11).AddCurrentValue CStr(TXTTo.text)
 xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtVal.text)
 xReport.ParameterFields(13).AddCurrentValue CStr(Me.XPMTxtRemarks.text)
 xReport.ParameterFields(14).AddCurrentValue CStr(LblValue.Caption)
 
 
 
 
xReport.EnableParameterPrompting = False
xReport.ApplicationName = App.Title
xReport.ReportAuthor = App.Title
Set CViewer = New ClsReportViewer
CViewer.FireReport xReport, WindowTarget, ""

RsData.Close
Set RsData = Nothing
Screen.MousePointer = vbDefault

End Function




Function print_report(Optional NoteSerial As String)
    
Dim MySQL As String
Dim RsData As New ADODB.Recordset
Dim xApp As New CRAXDRT.Application
Dim xReport As CRAXDRT.Report
Dim CViewer As ClsReportViewer
Dim StrReportTitle As String
Dim StrFileName As String
Dim Msg As String

MySQL = "Select * From Expanses_Order  where noteserial='" & NoteSerial & "'"

 
'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
'    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
'End If
'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
'    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
'End If

If SystemOptions.UserInterface = ArabicInterface Then
    StrFileName = App.path & "\Reports\" & "Expenses_order.rpt"
Else
    StrFileName = App.path & "\Reports\" & "Expenses_order.rpt"
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
    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
   ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
    StrReportTitle = "" '& StrAccountName
    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
    'End If
Else
 
    xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
     xReport.ParameterFields(4).AddCurrentValue get_branch_name(Val(my_branch))
    StrReportTitle = ""
    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
    'End If
End If
xReport.ParameterFields(3).AddCurrentValue user_name
xReport.reporttitle = StrReportTitle
xReport.EnableParameterPrompting = False
xReport.ApplicationName = App.Title
xReport.ReportAuthor = App.Title
Set CViewer = New ClsReportViewer
CViewer.FireReport xReport, WindowTarget, ""

RsData.Close
Set RsData = Nothing
Screen.MousePointer = vbDefault

End Function

Private Sub CmdHelp_Click()
SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdRemove_Click()
Dim x As Integer
    If SystemOptions.UserInterface = EnglishInterface Then
       x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
         x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

If x = vbNo Then Exit Sub
Dim sql As String

sql = "Delete  marakes_taklefa_temp where  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
    Cn.Execute sql, , adExecuteNoRecords
    
If CboPaymentType1.ListIndex = 0 Then
            If Fg_Journal.Rows > 1 Then
                If Fg_Journal.Rows = 2 Then
                    Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
                Else
                    If Me.Fg_Journal.Rows > 1 Then
                        If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                            Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                        End If
                    End If
                End If
            End If
            
            
            
            With Fg_Journal
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
             End With
ElseIf CboPaymentType1.ListIndex = 1 Then


            If VSFlexGrid1.Rows > 1 Then
                If VSFlexGrid1.Rows = 2 Then
                    Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                Else
                    If Me.VSFlexGrid1.Rows > 1 Then
                        If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                            Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                        End If
                    End If
                End If
            End If
            
            
            
            With VSFlexGrid1
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
             End With
             
             
 ElseIf CboPaymentType1.ListIndex = 2 Then


            If VSFlexGrid2.Rows > 1 Then
                If VSFlexGrid2.Rows = 2 Then
                    Me.VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
                Else
                    If Me.VSFlexGrid2.Rows > 1 Then
                        If Me.VSFlexGrid2.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                            Me.VSFlexGrid2.RemoveItem (Me.VSFlexGrid1.Row)
                        End If
                    End If
                End If
            End If
            
            
            
            With VSFlexGrid2
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
             End With
             
             
 Else
 
 Exit Sub
 End If

End Sub

Private Sub DCAccounts_Change()
If DCAccounts.BoundText = "" Then Exit Sub
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    DcboCreditSide.BoundText = DCAccounts.BoundText
End If

End Sub

Private Sub DCAccounts_Click(Area As Integer)
DCAccounts_Change
End Sub

Private Sub DcboBankName_Change()
'On Error Resume Next
If DcboBankName.BoundText = "" Then Exit Sub
Dim RsSavRec As ADODB.Recordset
Dim My_SQL As String
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
'    Me.DcboCreditSide.BoundText = "a2a3a2"
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

Set RsSavRec = New ADODB.Recordset
RsSavRec.CursorLocation = adUseClient
RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
  If SystemOptions.banks_Accounts3 = True Then
        Me.DcboCreditSide.BoundText = get_bank_Account(Val(Me.DcboBankName.BoundText), "Account_Code2")
        Else
        Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
    
   If CboPayMentType.ListIndex = 3 Or CboPayMentType.ListIndex = 5 Then
                     
 Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
End If

 'Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value



End If
End Sub

Private Sub DcboBox_Change()
If DcboBox.BoundText = "" Then Exit Sub
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
End If
End Sub



Private Sub DcboBox_Click(Area As Integer)
DcboBox_Change
End Sub

Private Sub dcBranch_Click(Area As Integer)
TxtSerial.text = ""
TxtSerial1.text = ""
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
CostCenterSearch.Show
CostCenterSearch.RetrunType = 3
End If
End Sub

Private Sub DCVendor_Change()
If DCVendor.BoundText = "" Then Exit Sub

If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", Val(Me.DCVendor.BoundText))
End If
Text2.text = Me.DCVendor.BoundText
End Sub

Private Sub DCVendor_Click(Area As Integer)
DCVendor_Change
End Sub

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, ByVal Col As Long)
 
Dim StrAccountCode As String
Dim Msg As String
Dim rs As New ADODB.Recordset
Dim StrSQL As String
Dim ClsAcc As New ClsAccounts
Dim LngRow As Long
With Fg_Journal
    Select Case .ColKey(Col)
       Case "ExpensesID"
              
             .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line

 
        Case "AccountName"
          '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
            StrAccountCode = .ComboData
            LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
             .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
             .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
             .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
             If SystemOptions.UserInterface = ArabicInterface Then
              StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
            Else
            StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
            End If
            
              rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
              If rs.RecordCount > 0 Then
              .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
              Else
              .TextMatrix(Row, .ColIndex("des")) = ""
              End If
              Case "value", "opr_fullcode"
           Dim sgl As String
                        Dim project_id As Integer
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), Val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                .TextMatrix(Row, .ColIndex("Value")) = 0
                End If
               
      Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
       sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
        Cn.Execute sgl, , adExecuteNoRecords
        
    '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    End Select
     Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    'to Add new row if needed
    If Row = .Rows - 1 Then
        .Rows = .Rows + 1
    End If
   ' ReLineGrid
End With
ReLineGrid
End Sub
Function calcnets()
If Me.CboPaymentType1.ListIndex = 0 Then
With Fg_Journal
Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
End With
ElseIf Me.CboPaymentType1.ListIndex = 1 Then
With Me.VSFlexGrid1
Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
End With
ElseIf Me.CboPaymentType1.ListIndex = 2 Then
With Me.VSFlexGrid2
Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
End With
End If
End Function
Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg_Journal
    If Row > .FixedRows Then
      '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
      '      Cancel = True
      '  End If
    End If
    Select Case .ColKey(Col)
        Case "value"
            .ComboList = ""
      Case "des"
        .ComboList = ""
          '  Cancel = True
            
         Case "Order_No"
        .ComboList = ""
    End Select
End With
End Sub

Private Sub Fg_Journal_DblClick()
  Exit Sub
  
Static lNoteRow&, lNoteCol&, r&, C&
With Fg_Journal
    ' clicking? no work
    'If Button <> 0 Then Exit Sub
    ' get mouse coordinates
    r = Fg_Journal.Row
    C = Fg_Journal.Col
    If Fg_Journal.ColKey(C) <> "Des" Then
        CboDes.Visible = False
        Exit Sub
    End If
    If Fg_Journal.TextMatrix(r, C) = "" Then
        'Exit Sub
    End If
    If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
        Exit Sub
    End If
    ' same cell or neighbour? no work
'    If r = lNoteRow And C = lNoteCol Then Exit Sub
'    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

    ' other cell, hide current note, if any
    If lNoteRow >= 0 And lNoteCol >= 0 Then
        Fg_Journal.SetFocus
        lNoteRow = -1
        lNoteCol = -1
    End If

    ' no note to show? then bail out
    If r <= 0 Or C <= 0 Then Exit Sub
    If TypeName(Fg_Journal.Cell(flexcpData, r, C)) <> "String" Then
        TxtDes.text = ""
    Else
        '
        TxtDes.text = Fg_Journal.Cell(flexcpData, r, C)
    End If
    ' show new note
    CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
    CboDes.Visible = True
    CboDes.ZOrder 0
    CboDes.SetFocus
    'save coordinates for next time
    lNoteRow = r
    lNoteCol = C
End With
End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, Shift As Integer)
With Fg_Journal
    Select Case .ColKey(.Col)
         Case "Order_No"
                           
                    If KeyCode = vbKeyF3 Then
                    Order_no_search.Show
                    Order_no_search.RetrunType = 4
                    End If
 Case "AccountName"
            If KeyCode = vbKeyF3 Then
                     FrmExpensesSearch.Show
                    FrmExpensesSearch.RetrunType = 2
             End If
 
 End Select
 End With

End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset

Dim StrSQL  As String
Dim StrAccountType As String
Dim StrComboList As String
Dim StrComboList1 As String

Dim Msg As String
'Case "DebitName"
'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
'Case "CreditName"
'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
With Fg_Journal
    Select Case .ColKey(Col)
         Case "AccountName"
                 StrSQL = "select * from Expenses_accounts"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
            If StrComboList <> "" Then
                StrComboList = "|" & StrComboList
            End If
            .ComboList = StrComboList
                  
          Case "opr_fullcode"
                    Dim project_id As Integer
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
          If SystemOptions.Items_or_operation = 1 Then
            StrSQL = "  select fullcode,name from terms_operations where project_id=" & project_id
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,name", "fullcode")
         ElseIf SystemOptions.Items_or_operation = 0 Then
           StrSQL = "  select fullcode,des from projects_des where project_id=" & project_id
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,des", "fullcode")
         
          End If
            If StrComboList1 <> "" Then
                StrComboList1 = "|" & StrComboList1
            End If
            .ComboList = StrComboList1
            
         
    End Select
End With

End Sub

Private Sub Form_Load()
Dim Dcombos As ClsDataCombos
Dim StrSQL As String

On Error GoTo ErrTrap

StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
fill_combo Me.DcCostCenter, StrSQL


ScreenNameArabic = "›« Ê—… ‘—«¡ «’· À«» "
ScreenNameEnglish = "F.A. Purchase Invoice"
RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1"


Set TTD = New clstooltipdemand
Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("FillData").Picture
Resize_Form Me
AddTip
SetDtpickerDate XPDtbTrans
Set Dcombos = New ClsDataCombos
Dcombos.GetBoxes Me.DcboBox
Dcombos.GetBanks Me.DcboBankName
Dcombos.GetUsers Me.DCboUserName
Dcombos.GetExpensesType XPCboExpensesType
Set cSearchDcbo = New clsDCboSearch
Set cSearchDcbo.Client = Me.XPCboExpensesType

Dcombos.GetAccountingCodes Me.DcboDebitSide
Dcombos.GetAccountingCodes Me.DcboCreditSide
 Dcombos.GetBranches Me.dcBranch
Dcombos.GetAccountingCodes Me.DCAccounts, True
With Me.CboPayMentType
    .Clear
    .AddItem "‰ﬁœÌ"
    .AddItem "‘Ìﬂ"
    .AddItem "«Ã·"
        .AddItem "ÕÊ«·…"
        .AddItem "Õ”«»"
                .AddItem "‘Ìﬂ „”œœ"

End With

With Me.CboPaymentType1
    .Clear
    .AddItem "„’«—Ì›"
    .AddItem "Õ”«»« "
    .AddItem "‘—«¡ «’· À«» "
End With

If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
End If


StrSQL = " select expanses_account,Project_name from projects  where not(expanses_account is null)"
fill_combo dcproject, StrSQL

'StrSQL = " select  CusID, CusName from TblCustemers  where Type=3"
If SystemOptions.UserInterface = ArabicInterface Then
StrSQL = " Select CusID,CusName From TblCustemers Where Type=2 or CustomerandVendor=1"
Else
StrSQL = " Select CusID,CusNamee From TblCustemers Where Type=2 or CustomerandVendor=1"
End If

fill_combo Me.DCVendor, StrSQL

Set rs = New ADODB.Recordset
StrSQL = "select * From notes_all where notetype=80 and bill_Type=2"
rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
XPBtnMove_Click 2
Me.TxtModFlg.text = "R"
If OPEN_NEW_SCREEN = True Then
Cmd_Click (0)
End If

Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrTrap
hide_logo = False
RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish
If rs.State = adStateOpen Then
    If Not (rs.EOF Or rs.BOF) Then
        If rs.EditMode <> adEditNone Then
            rs.CancelUpdate
        End If
    End If
    rs.Close
    Set rs = Nothing
End If
Set TTP = Nothing
'Set EmpReport = Nothing
TTD.Destroy
Exit Sub
ErrTrap:
End Sub

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, ByVal SpinningEnded As Boolean)
If ButtonID = vdsDownArrow Then
    If CboDes.IsDropped = False Then
        If PicHeight > 0 Then
            PicDes.Height = PicHeight
            PicDes.Width = PicWidth
        Else
            PicDes.Width = CboDes.Width - 10
            PicDes.Height = CboDes.Height * 8
        End If
        Debug.Print PicHeight
        Debug.Print PicWidth
        TxtDes.Visible = True
        TxtDes.text = Fg_Journal.Cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Des"))
        CboDes.DropDown PicDes.hWnd, vdsRightToLeft, vdsBottomToDown, vdsDownArrow, True, vdsSoftResize
        Debug.Print PicDes.Height & "Pic H " & "-----" & PicDes.Width & "Pic W"
    Else
        CboDes.CloseUp
    End If
End If
End Sub

Private Sub CboDes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{F4}"
End If
End Sub

Private Sub PicDes_Resize()
With PicDes
    LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
    TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
'    PicHeight = PicDes.Height
'    PicWidth = PicDes.Width
End With

End Sub

 

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Order_no_search.Show
Order_no_search.RetrunType = 1
End If
End Sub

Private Sub TxtDes_LostFocus()
PicHeight = PicDes.Height
PicWidth = PicDes.Width
CboDes.CloseUp
CboDes.Visible = False
End Sub

Private Sub TxtDes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    PutData
    CboDes.CloseUp
End If
End Sub

Private Sub TxtModFlg_Change()
On Error GoTo ErrTrap
Select Case Me.TxtModFlg.text
    Case "R"
 
        
         Me.VSFlexGrid1.Enabled = False
        Me.Fg_Journal.Enabled = False
        Frame1.Enabled = False

        
        Me.Cmd(2).Enabled = False
        Me.Cmd(3).Enabled = False
        CmdRemove.Enabled = False
        Me.Cmd(0).Enabled = True
        Me.Cmd(1).Enabled = True
        Me.Cmd(4).Enabled = True
        Me.Cmd(5).Enabled = True
        Me.XPBtnMove(0).Enabled = True
        Me.XPBtnMove(1).Enabled = True
        Me.XPBtnMove(2).Enabled = True
        Me.XPBtnMove(3).Enabled = True
        
        XPTxtVal.Locked = True
'        XPCboProfLevel.Locked = True
'        XPTxtProfMail.Locked = True
'        XPTxtPhone.Locked = True
'        XPTxtMobile.Locked = True
        XPMTxtRemarks.Locked = True
        XPCboExpensesType.Locked = True
        Me.DcboBox.Locked = True
        XPDtbTrans.Enabled = False
        If rs.RecordCount < 1 Then
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            
        End If
    Case "N"
 
        
        
  Me.VSFlexGrid1.Enabled = True
        Me.Fg_Journal.Enabled = True
        Frame1.Enabled = True


        Me.Cmd(2).Enabled = True
        Me.Cmd(3).Enabled = True
        CmdRemove.Enabled = True
        Me.Cmd(0).Enabled = False
        Me.Cmd(1).Enabled = False
        Me.Cmd(4).Enabled = False
        Me.Cmd(5).Enabled = False
     '   Me.XPBtnMove(0).Enabled = False
     '   Me.XPBtnMove(1).Enabled = False
     '   Me.XPBtnMove(2).Enabled = False
     '   Me.XPBtnMove(3).Enabled = False
        
        XPTxtVal.Locked = False
'        XPCboProfLevel.Locked = False
'        XPTxtProfMail.Locked = False
'        XPTxtPhone.Locked = False
'        XPTxtMobile.Locked = False
        XPMTxtRemarks.Locked = False
        XPCboExpensesType.Locked = False
        Me.DcboBox.Locked = False
        XPDtbTrans.Enabled = True
        XPDtbTrans.value = Date
    Case "E"
 
        
               Me.VSFlexGrid1.Enabled = True
        Me.Fg_Journal.Enabled = True
       Frame1.Enabled = True
       
       
        CmdRemove.Enabled = True
        Me.Cmd(2).Enabled = True
        Me.Cmd(3).Enabled = True
        
        Me.Cmd(0).Enabled = False
        Me.Cmd(1).Enabled = False
        Me.Cmd(4).Enabled = False
        Me.Cmd(5).Enabled = False
        Me.XPBtnMove(0).Enabled = False
        Me.XPBtnMove(1).Enabled = False
        Me.XPBtnMove(2).Enabled = False
        Me.XPBtnMove(3).Enabled = False
        
        XPTxtVal.Locked = False
'        XPCboProfLevel.Locked = False
'        XPTxtProfMail.Locked = False
'        XPTxtPhone.Locked = False
'        XPTxtMobile.Locked = False
        XPMTxtRemarks.Locked = False
        XPCboExpensesType.Locked = False
        Me.DcboBox.Locked = False
        XPDtbTrans.Enabled = True
End Select
Exit Sub
ErrTrap:
End Sub

Public Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'check_cost_center
Dim StrAccountCode As String
Dim Msg As String
Dim rs As New ADODB.Recordset
Dim StrSQL As String
Dim ClsAcc As New ClsAccounts
Dim LngRow As Long
 Dim sql As String
 
With VSFlexGrid1
    Select Case .ColKey(Col)
    
    Case "Value"
    Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
        Case "DebitValue", "CreditValue"
 


'remove destribution
     
       ' sgl = "update  marakes_taklefa_temp  set value=0 where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
       ' Cn.Execute sgl, , adExecuteNoRecords
    
    
    
            .TextMatrix(Row, Col) = Val(.TextMatrix(Row, Col))
            
            If .ColKey(Col) = "DebitValue" Then
                .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                ' Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                 
                '    Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
            ElseIf .ColKey(Col) = "CreditValue" Then
                 .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                 .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                ' Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                '     Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
            End If
            .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
            .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
            
        Case "DebitValueE", "CreditValueE"
             .TextMatrix(Row, Col) = Val(.TextMatrix(Row, Col))

            If .ColKey(Col) = "DebitValueE" Then
               .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
                .TextMatrix(Row, .ColIndex("CreditValue")) = 0

                If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                Else
                 .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE"))
                 End If
               '
               '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
               '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
               '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
               '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
            ElseIf .ColKey(Col) = "CreditValueE" Then
                .Cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                 .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                 .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                 .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                 Else
                 .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE"))
                 End If
                 
               '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
               '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
               '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
               '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
            End If
            
        Case "Account_Serial"
         .TextMatrix(Row, .ColIndex("userid")) = user_id
            .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))
            If .TextMatrix(Row, Col) = "" Then
                Exit Sub
            End If
            StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & _
            "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & _
            "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & _
            " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
            Set rs = Nothing
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rs.BOF Or rs.EOF) Then
                If BolEditOnMainAccounts = False Then
                    'If LastAccount(rs("Account_Code").value) = False Then
                    '    .TextMatrix(Row, Col) = ""
                    '    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    '    Exit Sub
                    'End If
                End If
                .TextMatrix(Row, .ColIndex("AccountCode")) = _
                    IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Row, .ColIndex("AccountName")) = _
                    IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                    .TextMatrix(Row, .ColIndex("AccountName")) = _
                    IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    
                    
                    
                    End If
                    
                     .TextMatrix(Row, .ColIndex("cost_center")) = _
                    IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)
                    
                    
                    
                    
Dim rs2 As ADODB.Recordset
Dim My_SQL As String
If IsNull(rs("currenct_code").value) Then

                    .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                    
                    
                    
                    .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                    
GoTo xx
End If

    My_SQL = "  select * from currency WHERE id=" & Val(rs("currenct_code").value)

Set rs2 = New ADODB.Recordset
rs2.CursorLocation = adUseClient
rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 

 
 
 
                    .TextMatrix(Row, .ColIndex("currenct_code")) = _
                    IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    
                    .TextMatrix(Row, .ColIndex("rate")) = _
                    IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
xx:
            Else
                GetMsgs 130, vbExclamation
                .TextMatrix(Row, Col) = ""
                .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                Exit Sub
            End If
            rs.Close
            Set rs = Nothing
        Case "AccountName"
        
        'sgl = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
        'Cn.Execute sgl, , adExecuteNoRecords
    
    
                        .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
            StrAccountCode = .ComboData
            LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
            If LngRow <> -1 Then
                'Msg = "Â–« «·Õ”«» „ÊÃÊœ „”»ﬁ«  ›Ï «·”ÿ— " & .TextMatrix(LngRow, .ColIndex("LineNo"))
                'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                '.TextMatrix(Row, Col) = ""
                '.TextMatrix(Row, .ColIndex("AccountCode")) = ""
                'Exit Sub
            End If
            Set ClsAcc = New ClsAccounts
            If BolEditOnMainAccounts = False Then
                'If LastAccount(StrAccountCode) = False Then
                '    .TextMatrix(Row, Col) = ""
                '    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                'Else

                   .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                   .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                'End If
            Else
                 .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
 
                 .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
            End If
            Set ClsAcc = Nothing
            
            
                 StrSQL = "SELECT ACCOUNTS.cost_center ,ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & _
            "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & _
            "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & _
            " From ACCOUNTS Where ACCOUNTS.Account_Name='" & Trim(.TextMatrix(Row, Col)) & "'"
            Set rs = Nothing
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rs.BOF Or rs.EOF) Then
                  .TextMatrix(Row, .ColIndex("cost_center")) = _
                    IIf(IsNull(rs("cost_center").value), vbFalse, rs("cost_center").value)
            
            
'Dim rs2 As ADODB.Recordset
'Dim My_SQL As String
If IsNull(rs("currenct_code").value) Then
       .TextMatrix(Row, .ColIndex("currenct_code")) = ""
        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
GoTo ll
End If
    My_SQL = "  select * from currency WHERE id=" & rs("currenct_code").value

Set rs2 = New ADODB.Recordset
rs2.CursorLocation = adUseClient
rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 


         .TextMatrix(Row, .ColIndex("currenct_code")) = _
                    IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = _
                    IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value)
ll:
End If


    End Select
    'to Add new row if needed
    If Row = .Rows - 1 Then
        .Rows = .Rows + 1
    End If
    ReLineGrid
    
 


End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid1
    If Row > .FixedRows Then
      '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
      '      Cancel = True
      '  End If
    End If
    Select Case .ColKey(Col)
        Case "Value"
            .ComboList = ""
      Case "Account_Serial"
        .ComboList = ""
          '  Cancel = True
            
    End Select
End With

End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Account_search.Show
Account_search.case_id = 80

End If
End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
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
         Case "AccountName"
        
                 
                
                'Full Path Display
                  If SystemOptions.UserInterface = EnglishInterface Then
                
                StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & _
                "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & _
                " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & _
                "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & _
                "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & _
                "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
             '   If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
             '   End If
                If OptSort(1).value = True Then
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                Else
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                End If
                
                Else
                
                                StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & _
                "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & _
                " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & _
                "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & _
                "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & _
                "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
           '     If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If
           '     End If
                If OptSort(1).value = True Then
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                Else
                    StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                End If
                
                End If
                
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
                Debug.Print StrSQL
 
            If StrComboList <> "" Then
                StrComboList = "|" & StrComboList
            End If
            .ComboList = StrComboList
    End Select
End With

End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
 
Dim StrAccountCode As String
Dim Msg As String
Dim rs As New ADODB.Recordset
Dim StrSQL As String
Dim ClsAcc As New ClsAccounts
Dim LngRow As Long
With VSFlexGrid2
    Select Case .ColKey(Col)
 

 
        Case "AccountName"
          '  .TextMatrix(Row, .ColIndex("userid")) = user_id
             Dim groupid As Integer
             Dim branch_id As Integer
            StrAccountCode = .ComboData
            LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
             .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
            
              StrSQL = "select * from FixedAssets where id=" & Val(StrAccountCode)
           
            
              rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
              If rs.RecordCount > 0 Then
              groupid = IIf(IsNull(rs("group_id").value), "", rs("group_id").value)
              .TextMatrix(Row, .ColIndex("groupid")) = groupid
              branch_id = IIf(IsNull(rs("Branch_NO").value), "", rs("Branch_NO").value)
              .TextMatrix(Row, .ColIndex("branch_id")) = branch_id
              
              
              
              Else
              .TextMatrix(Row, .ColIndex("groupid")) = 0
              groupid = 0
              branch_id = 0
              .TextMatrix(Row, .ColIndex("branch_id")) = 0
              End If
              
             .TextMatrix(Row, .ColIndex("AccountCode")) = get_FixedAsset_Account(groupid, branch_id)
      
        
               
      Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    
        
    '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    End Select
     Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
    'to Add new row if needed
    If Row = .Rows - 1 Then
        .Rows = .Rows + 1
    End If
   ' ReLineGrid
End With
With Me.VSFlexGrid2
If Me.TxtModFlg <> "E" Then Exit Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
If Col = .ColIndex("AccountName") Then
        LogTextA = "   ⁄œÌ· «·«’· «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        LogTextE = "  Change F.A. To " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
ElseIf Col = .ColIndex("value") Then
        LogTextA = "   ⁄œÌ· «·ﬁÌ„…  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("value")) & " ··«’·   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        LogTextE = "  Change value" & .Cell(flexcpTextDisplay, Row, .ColIndex("value")) & " To F.A. " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
 ElseIf Col = .ColIndex("des") Then
        LogTextA = "   ⁄œÌ· «·‘—Õ  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("des")) & " ··«’·   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
        LogTextE = "  Change Des " & .Cell(flexcpTextDisplay, Row, .ColIndex("des")) & " To  F.A. " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
End If
AddToLogFile CInt(user_id), 300, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
End With

ReLineGrid


End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid2
 
    Select Case .ColKey(Col)
        Case "value"
            .ComboList = ""
      Case "des"
        .ComboList = ""
    
    End Select
End With

End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset

Dim StrSQL  As String
Dim StrAccountType As String
Dim StrComboList As String
Dim StrComboList1 As String

Dim Msg As String
'Case "DebitName"
'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
'Case "CreditName"
'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
With VSFlexGrid2
    Select Case .ColKey(Col)
         Case "AccountName"
                 StrSQL = "select * from FixedAssets where New_or_opening=0 and PurchasePrice=0 order by Name"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Name", "Id")
            If StrComboList <> "" Then
                StrComboList = "|" & StrComboList
            End If
            .ComboList = StrComboList
                  
             
         
    End Select
End With

End Sub

Private Sub XPBtnMove_Click(Index As Integer)
On Error GoTo ErrTrap
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
Retrive
Exit Sub
ErrTrap:
End Sub
Public Sub Retrive(Optional Lngid As String = "")
Dim RsDev As ADODB.Recordset
Dim StrSQL As String
Dim i As Integer

On Error GoTo ErrTrap
          Fg_Journal.Clear flexClearScrollable, flexClearEverything
          Fg_Journal.Rows = 3
          
          VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
          VSFlexGrid1.Rows = 2
          
          VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
          VSFlexGrid2.Rows = 2
If rs.RecordCount < 1 Then
    XPTxtCurrent.Caption = 0
    XPTxtCount.Caption = 0
    Exit Sub
End If
If rs.EOF Or rs.BOF Then
    Exit Sub
End If
If Lngid = "-1" Then
Cmd_Click (0)
End If

If Lngid <> "" Then
  '  If XPTxtID.text <> 0 Then
        rs.Find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst
        If rs.EOF Or rs.BOF Then
         clear_all Me
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "›« Ê—… €Ì— „”Ã·… ", vbInformation
         Else
         MsgBox " Un Refistered Bill ", vbInformation
         End If
            Exit Sub
        End If
  '  End If
End If


 If Not IsNull(rs("general_cost_center").value) Then
    Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
   Else
    Me.DcCostCenter.BoundText = ""
 End If
Me.Text1.text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
Me.TXT_order_no.text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)

 TXT_A_NoteID.text = IIf(IsNull(rs("A_NoteID").value), "", Val(rs("A_NoteID").value))

XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", Val(rs("NoteID").value))
Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
TXTTo.text = IIf(IsNull(rs("too").value), "", rs("too").value)
txt_general_des.text = IIf(IsNull(rs("general_des").value), "", rs("general_des").value)

XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
XPCboExpensesType.BoundText = IIf(IsNull(rs("ExpensesID").value), "", rs("ExpensesID").value)


If (rs("bill_Type").value) = 0 Then
 Me.CboPaymentType1.ListIndex = 0
ElseIf (rs("bill_Type").value) = 1 Then
Me.CboPaymentType1.ListIndex = 1
ElseIf (rs("bill_Type").value) = 2 Then
Me.CboPaymentType1.ListIndex = 2

End If

CboPaymentType1_Change

If IsNull(rs("NoteCashingType").value) Then
    Me.CboPayMentType.ListIndex = 0
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
    Me.DcboBankName.BoundText = ""
    Me.TxtChequeNumber.text = ""
    DCVendor.BoundText = ""
ElseIf rs("NoteCashingType").value = 0 Then
    Me.CboPayMentType.ListIndex = 0
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
    Me.DcboBankName.BoundText = ""
    Me.TxtChequeNumber.text = ""
    DCVendor.BoundText = ""
ElseIf rs("NoteCashingType").value = 1 Then
    Me.CboPayMentType.ListIndex = 1
    Me.DcboBox.BoundText = ""
    Me.DcboBankName.BoundText = rs("BankID").value
    Me.TxtChequeNumber.text = rs("ChqueNum").value
    Me.DtpChequeDueDate.value = rs("DueDate").value
    DCVendor.BoundText = ""
ElseIf rs("NoteCashingType").value = 2 Then
    Me.CboPayMentType.ListIndex = 2
    Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = ""
    Me.TxtChequeNumber.text = ""
    
    Me.DCVendor.BoundText = rs("CusID").value

ElseIf rs("NoteCashingType").value = 3 Then
    Me.CboPayMentType.ListIndex = 3
    Me.DcboBox.BoundText = ""
    Me.DcboBankName.BoundText = rs("BankID").value
    Me.TxtChequeNumber.text = rs("ChqueNum").value
    Me.DtpChequeDueDate.value = rs("DueDate").value
    DCVendor.BoundText = ""
ElseIf rs("NoteCashingType").value = 5 Then
    Me.CboPayMentType.ListIndex = 5
    Me.DcboBox.BoundText = ""
    Me.DcboBankName.BoundText = rs("BankID").value
    Me.TxtChequeNumber.text = rs("ChqueNum").value
    Me.DtpChequeDueDate.value = rs("DueDate").value
    DCVendor.BoundText = ""
    
ElseIf rs("NoteCashingType").value = 4 Then
    Me.CboPayMentType.ListIndex = 4
    Me.DCAccounts.BoundText = IIf(IsNull(rs("AccountCode").value), "", rs("AccountCode").value)
   DcboBox.BoundText = ""
    Me.DcboBankName.BoundText = ""
    Me.TxtChequeNumber.text = ""
    DCVendor.BoundText = ""
    
End If
CboPayMentType_Change

'ÿMe.DcboBox.BoundText = IIf(IsNull(Rs("BoxID").value), "", Rs("BoxID").value)
'DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))

      If rs("NoteCashingType").value = 0 Then
                DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
       ElseIf rs("NoteCashingType").value = 1 Then
               DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", Val(Me.DcboBankName.BoundText))
       ElseIf rs("NoteCashingType").value = 2 Then
               DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", Val(Me.DCVendor.BoundText))
     End If
     
            

Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
Me.Txt_Numorder.text = IIf(IsNull(rs("NumOrderInpot").value), "", rs("NumOrderInpot").value)
Me.TxtSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
Me.TxtSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)

Me.dcproject.BoundText = IIf(IsNull(rs("project_Expensen_account").value), "", rs("project_Expensen_account").value)


If CboPaymentType1.ListIndex = 1 Then 'Õ”«Ì« 


StrSQL = "SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, "
StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.UserID , dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[value],dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description"
StrSQL = StrSQL + " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
StrSQL = StrSQL + " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = " & Val(rs("A_NoteID").value) & ")"
StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"



     Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If RsDev.RecordCount > 0 Then
      RsDev.MoveFirst
    End If
    
    With Me.VSFlexGrid1
 
    .Rows = .FixedRows + RsDev.RecordCount
 
    For i = .FixedRows To .Rows
      .TextMatrix(i, .ColIndex("LineNo")) = i
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), _
            "", RsDev("Account_Code").value)
            
            
           .TextMatrix(i, .ColIndex("account_serial")) = IIf(IsNull(RsDev("account_serial").value), _
            "", RsDev("account_serial").value)
            
            
            If SystemOptions.UserInterface = ArabicInterface Then
            
                 .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), _
            "", RsDev("Account_Name").value)
            Else
               .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), _
            "", RsDev("Account_NameEng").value)
            End If
            
 
        
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), _
            "", RsDev("Value").value)
            
                .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), _
            "", RsDev("Double_Entry_Vouchers_Description").value)
            

            
        RsDev.MoveNext
    Next i
    
    End With


Exit Sub
End If
'-----------------------------------------------------------------------------
If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then '«·«’Ê·
 '   StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(Me.XPTxtID.text)
 '   StrSQL = StrSQL + " Order By DEV_ID_Line_No "
' StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.*,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name FROM    dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code WHERE     dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID =" & Val(Me.XPTxtID.text) & "Order By DEV_ID_Line_No"

'StrSQL = "SELECT   dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode,   dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit , dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID ,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description  FROM         dbo.ACCOUNTS INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
'StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0  and dbo.DOUBLE_ENTREY_VOUCHERS.notes_all =" & Val(Me.XPTxtID.text) & ") "
'StrSQL = StrSQL + "ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
StrSQL = "SELECT  dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetbranch_id , dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetgroupid, dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetID ,  dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,"
StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name,"
StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID AS Expr1,"
StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description , dbo.Notes.order_no"
StrSQL = StrSQL + " FROM         dbo.ACCOUNTS INNER JOIN"
StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
StrSQL = StrSQL + " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & Val(Me.XPTxtID.text) & ")"
StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsDev.BOF Or rs.EOF) Then
        Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
        Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
        RsDev.MoveFirst
        For i = 1 To RsDev.RecordCount
            If RsDev("Credit_Or_Debit").value = 0 Then
                Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
            ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
            End If
            RsDev.MoveNext
        Next i
    
    
      RsDev.MoveFirst
    
    With Me.VSFlexGrid2
    If Me.dcproject.BoundText = "" Then
   .Rows = .FixedRows + RsDev.RecordCount
    Else
    .Rows = .FixedRows + RsDev.RecordCount - 1
    End If
    For i = .FixedRows To .Rows - 1
        .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("DEV_ID_Line_No").value), _
            "", RsDev("DEV_ID_Line_No").value)
            
            
                  .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), _
            "", RsDev("DEV_ID_Line_No1").value)
            
            .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("FixedAssetId").value), _
            "", RsDev("FixedAssetId").value)
            
            
            
       .TextMatrix(i, .ColIndex("AccountName")) = getFixedAsstName(Val(.TextMatrix(i, .ColIndex("id"))), "name")
           
            .TextMatrix(i, .ColIndex("groupid")) = IIf(IsNull(RsDev("FixedAssetgroupid").value), _
            "", RsDev("FixedAssetgroupid").value)
            
            .TextMatrix(i, .ColIndex("branch_id")) = IIf(IsNull(RsDev("FixedAssetbranch_id").value), _
            "", RsDev("FixedAssetbranch_id").value)
                    
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), _
            "", RsDev("Account_Code").value)
             
       
        .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), _
            "", RsDev("Double_Entry_Vouchers_Description").value)
  

        
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), _
            "", RsDev("Value").value)
     
            
 
        RsDev.MoveNext
    Next i
        Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
         ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
  '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), _
  '  .Rows - 1, .ColIndex("CreditValue"))
  '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), _
  '  .Rows - 1, .ColIndex("DebitValue"))
End With
End If

End If
'-----------------------------------------------------------------------------
XPTxtCurrent.Caption = rs.AbsolutePosition
XPTxtCount.Caption = rs.RecordCount
ReLineGrid
Exit Sub
ErrTrap:
End Sub
Private Sub SaveData()
Dim Msg As String
Dim RsTemp As New ADODB.Recordset
Dim StrSQL As String
Dim BeginTrans As Boolean
Dim RsDev As ADODB.Recordset
Dim LngDevID As Long

On Error GoTo ErrTrap
If Me.TxtModFlg.text <> "R" Then

   If Me.CboPaymentType1.ListIndex = -1 Then
   If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·›« Ê—… ...!!!"
    Else
    Msg = "Select Bill Type ...!!!"
    End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboPayMentType.SetFocus
        Exit Sub
    End If
    
    
   If Me.CboPayMentType.ListIndex = -1 Then
   If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
    Else
    Msg = "Select Payment method ...!!!"
    End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboPayMentType.SetFocus
        Exit Sub
    End If
    
       If Me.CboPayMentType.ListIndex = 2 Then
       If Trim(Me.DCVendor.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·„Ê—œ..!!"
                Else
                Msg = "Select vendor..!!"
                End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCVendor.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
        End If
        
       If Me.CboPayMentType.ListIndex = 4 Then
       If Trim(Me.DCAccounts.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— ·Õ”«»..!!"
                Else
                Msg = "Select Account..!!"
                End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCAccounts.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
        
        End If
        
    
    If Me.CboPayMentType.ListIndex = 0 Then
        If Trim(Me.DcboBox.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
        Else
        Msg = "Select Box..!!"
        End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboBox.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
    ElseIf Me.CboPayMentType.ListIndex = 1 Then
                                                                If Me.DcboBankName.BoundText = "" Then
                                                                If SystemOptions.UserInterface = ArabicInterface Then
                                                                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                                                                Else
                                                                Msg = "Select Bank...!!"
                                                                
                                                                End If
                                                                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                                                    DcboBankName.SetFocus
                                                                    SendKeys "{F4}"
                                                                    Exit Sub
                                                                End If
                    If Trim$(Me.TxtChequeNumber.text) = "" Then
                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                                            Else
                                            Msg = "Enter Cheque No:...!!"
                                            End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        TxtChequeNumber.SetFocus
                        Exit Sub
                    End If
                                                                '        If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
                                                    '                        If SystemOptions.UserInterface = ArabicInterface Then
                                                    '                            Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
                                                    '                        Else
                                                    '                        Msg = "Cheque Due Date Not Valid...!!"
                                                    '
                                                    '                        End If
                                                    '            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                                    '            DtpChequeDueDate.SetFocus
                                                    '            SendKeys "{F4}"
                                                    '            Exit Sub
                                                    '        End If
             ElseIf Me.CboPayMentType.ListIndex = 3 Then
                        If Me.DcboBankName.BoundText = "" Then
                                        If SystemOptions.UserInterface = ArabicInterface Then
                                            Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                                        Else
                                        Msg = "Select Bank...!!"
                                        
                                        End If
                            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            DcboBankName.SetFocus
                            SendKeys "{F4}"
                            Exit Sub
                        End If
                                                If Trim$(Me.TxtChequeNumber.text) = "" Then
                                                                If SystemOptions.UserInterface = ArabicInterface Then
                                                                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·ÕÊ«·…...!!"
                                                                Else
                                                                Msg = "Enter Cheque No:...!!"
                                                                End If
                                                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                                    TxtChequeNumber.SetFocus
                                                    Exit Sub
                                                    
                                                End If
             ElseIf Me.CboPayMentType.ListIndex = 5 Then
                        If Me.DcboBankName.BoundText = "" Then
                                        If SystemOptions.UserInterface = ArabicInterface Then
                                            Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                                        Else
                                        Msg = "Select Bank...!!"
                                        
                                        End If
                            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            DcboBankName.SetFocus
                            SendKeys "{F4}"
                            Exit Sub
                        End If
                                                If Trim$(Me.TxtChequeNumber.text) = "" Then
                                                                If SystemOptions.UserInterface = ArabicInterface Then
                                                                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                                                                Else
                                                                Msg = "Enter Cheque No:...!!"
                                                                End If
                                                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                                                    TxtChequeNumber.SetFocus
                                                    Exit Sub
                                                    
                                                End If
                                                
       
       
       End If
       
 
    
    If Me.TxtModFlg.text = "N" Then
        If Me.CboPayMentType.ListIndex = 0 Then
            If Val(Me.DcboBox.BoundText) <> 0 Then
                If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.XPTxtVal.text), _
                    XPDtbTrans.value) = False Then
                    Exit Sub
                End If
            End If
        End If
    ElseIf Me.TxtModFlg.text = "E" Then
        If Me.CboPayMentType.ListIndex = 0 Then
            If Val(Me.DcboBox.BoundText) <> 0 Then
                If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.XPTxtVal.text), _
                    XPDtbTrans.value, , , Val(Me.XPTxtID.text)) = False Then
                    Exit Sub
                End If
            End If
        End If
    End If
    
     Dim xrow As Integer
With Fg_Journal
For xrow = .Rows - 1 To 2 Step -1
If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

 .Rows = .Rows - 1
End If
Next xrow
End With


    
With Me.VSFlexGrid1
For xrow = .Rows - 1 To 2 Step -1
If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

 .Rows = .Rows - 1
End If
Next xrow
End With

With Me.VSFlexGrid2
For xrow = .Rows - 1 To 2 Step -1
If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

 .Rows = .Rows - 1
End If
Next xrow
End With

 Dim i As Integer
If CboPaymentType1.ListIndex = 2 Then
  With Me.VSFlexGrid2
     For i = .FixedRows To .Rows - 1
         If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
              '////////////////////////////////////////notes
                
               
               If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox "·«Ì ÌÊÃœ «’· ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
               Else
               MsgBox "Select FixedAsset in line no" & i, vbCritical
               End If
               Exit Sub
              
        End If
        
    Next i



End With



   With VSFlexGrid2
     For i = .FixedRows To .Rows - 1
         If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Or Val(.TextMatrix(i, .ColIndex("value"))) <= 0 Then
              '////////////////////////////////////////notes
                
               
               If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox "·«Ì ÌÊÃœ ﬁÌ„… ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
               Else
               MsgBox "Enter Value in line no" & i, vbCritical
               End If
               
              Exit Sub
        End If
        
    Next i
End With
  Dim noOfInstallments As Integer 'Â–« «·Ã“¡ Ì √ﬂœ „‰  ‰›Ì– «ﬁ”«ÿ «Â·«ﬂ
  Dim msgStr As String
  With Me.VSFlexGrid2
     For i = .FixedRows To .Rows - 1
         If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
              '////////////////////////////////////////notes
                
               noOfInstallments = CheCkInstallmentCount(Val(.TextMatrix(i, .ColIndex("id"))))
               If noOfInstallments > 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                        msgStr = " ·« Ì„ﬂ‰ «· ⁄œÌ·  „  ‰›Ì– «ﬁ”«ÿ ⁄·Ï «·«’·  " & Chr(13)
                        msgStr = msgStr & .TextMatrix(i, .ColIndex("AccountName")) & Chr(13)
                        msgStr = msgStr & "⁄œœ «·«ﬁ”«ÿ «·„‰›–… Õ Ï «·«‰ " & noOfInstallments
                        MsgBox msgStr, vbCritical
                        Else
                           msgStr = " Can't Edit Fixed Asset   " & Chr(13)
                        msgStr = msgStr & .TextMatrix(i, .ColIndex("AccountName")) & Chr(13)
                        msgStr = msgStr & "No Of Executed Installments " & noOfInstallments
                        MsgBox msgStr, vbCritical
                        End If
                        Exit Sub
               End If
               
               
              
        End If
        
    Next i
End With





End If

If CboPaymentType1.ListIndex = 0 Then
  With Fg_Journal
     For i = .FixedRows To .Rows - 1
         If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
              '////////////////////////////////////////notes
                
               
               If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox "·«Ì ÌÊÃœ „’—Ê› ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
               Else
               MsgBox "Select Expenses in line no" & i, vbCritical
               End If
               Exit Sub
              
        End If
        
    Next i
End With
   With Fg_Journal
     For i = .FixedRows To .Rows - 1
         If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Or Val(.TextMatrix(i, .ColIndex("value"))) <= 0 Then
              '////////////////////////////////////////notes
                
               
               If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox "·«Ì ÌÊÃœ ﬁÌ„… ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
               Else
               MsgBox "Enter Value in line no" & i, vbCritical
               End If
               
              Exit Sub
        End If
        
    Next i
End With
End If
calcnets

    '-------------------------------------------------------------------------------------------
 
    '-------------------------------------------------------------------------------------------
    If TxtSerial.text = "" Then
         If Notes_coding(Val(my_branch), XPDtbTrans.value) = "error" Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
         Else
         MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number ": Exit Sub
         End If
         Else
         
         If Notes_coding(Val(my_branch), XPDtbTrans.value) = "" Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
         Else
         MsgBox "You must Define JE Coding ": Exit Sub
         End If
         Else
         TxtSerial.text = Notes_coding(Val(my_branch), XPDtbTrans.value)
         End If
         End If
    End If
 
      If TxtSerial1.text = "" Then
         If Voucher_coding(Val(my_branch), XPDtbTrans.value, 8, 80) = "error" Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
         Else
         MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
         End If
         Else
         
         If Voucher_coding(Val(my_branch), XPDtbTrans.value, 8, 80) = "" Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
         Else
         MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
         End If
         Else
         TxtSerial1.text = Voucher_coding(Val(my_branch), XPDtbTrans.value, 8, 80)
         End If
         End If
    End If
    
    
    Cn.BeginTrans
    BeginTrans = True
    
    '///////////////NOTESALL
        Dim A_NoteID As Long
    If TxtModFlg.text = "N" Then
        XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
        Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=80"))
        rs.AddNew
      
    ElseIf Me.TxtModFlg.text = "E" Then
    
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where notes_all=" & Val(XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        StrSQL = "Delete From notes Where notes_all=" & Val(XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
       
       If DcCostCenter.BoundText <> "" Then
        StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        End If
        
        
        
    End If
    
  '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
  
    rs("NoteID").value = Val(XPTxtID.text)
    
    rs("bill_Type").value = Me.CboPaymentType1.ListIndex
    rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
    rs("foxy_no").value = Val(Text1.text)
    rs("order_no").value = TXT_order_no.text
        rs("branch_no").value = Val(Me.dcBranch.BoundText)
    rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, XPTxtVal.text)
    rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
    rs("too").value = IIf(TXTTo.text = "", "", Trim(TXTTo.text))
    rs("general_des").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    
    
    rs("CusID").value = Null
    rs("NoteType").value = 80
    rs("NoteDate").value = XPDtbTrans.value
    rs("UserID").value = user_id
    rs("ExpensesID").value = IIf(XPCboExpensesType.text = "", Null, XPCboExpensesType.BoundText)
  
Dim bankDes As String
    If Me.CboPayMentType.ListIndex = 0 Then
        rs("BoxID").value = Val(DcboBox.BoundText)
        rs("BankID").value = Null
        rs("ChqueNum").value = Null
        rs("DueDate").value = Null
        rs("NoteCashingType").value = 0
    ElseIf Me.CboPayMentType.ListIndex = 1 Then
        rs("BoxID").value = Null
        rs("BankID").value = Val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("NoteCashingType").value = 1
                If SystemOptions.UserInterface = ArabicInterface Then
        bankDes = "  ’—› »‘Ìﬂ —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text & "»‰«¡ ⁄·Ï" & txt_general_des.text
        Else
        bankDes = "  Check No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text & "Base ON  " & txt_general_des.text
        
        End If
        
    ElseIf Me.CboPayMentType.ListIndex = 2 Then
        rs("NoteCashingType").value = 2
        rs("CusID").value = Val(Me.DCVendor.BoundText)
    ElseIf Me.CboPayMentType.ListIndex = 3 Then
        rs("BoxID").value = Null
        rs("BankID").value = Val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("NoteCashingType").value = 3
                If SystemOptions.UserInterface = ArabicInterface Then
        bankDes = "  ’—› »ÕÊ«·…  —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
        Else
        bankDes = "  Bank Transfere No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
        End If
    
    ElseIf Me.CboPayMentType.ListIndex = 5 Then
        rs("BoxID").value = Null
        rs("BankID").value = Val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("NoteCashingType").value = 5
                If SystemOptions.UserInterface = ArabicInterface Then
        bankDes = "  ’—› »‘Ìﬂ „”œœ  —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
        Else
        bankDes = "  Bank Transfere No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
        End If
        
    ElseIf Me.CboPayMentType.ListIndex = 4 Then
        rs("BoxID").value = Null
        rs("BankID").value = Null
        rs("ChqueNum").value = Null
        rs("DueDate").value = Null
        rs("NoteCashingType").value = 4
        rs("AccountCode").value = (Me.DCAccounts.BoundText)
        
             If SystemOptions.UserInterface = ArabicInterface Then
        bankDes = txt_general_des.text
        Else
        bankDes = txt_general_des.text
        
        End If
       
       
    End If
    
    rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
    rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    rs("Buy").value = "0"
    rs("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    rs("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
    rs("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”·   ›« Ê—…
    rs("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    rs("numbering_type1").value = sand_numbering_type(8) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
     
    rs("sanad_year").value = year(XPDtbTrans.value)
    rs("sanad_month").value = Month(XPDtbTrans.value)
    If dcproject.BoundText <> "" Then
    rs("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
    Else
    rs("note_value_by_characters").value = WriteNo(Format(Val(Me.XPTxtVal.text) * 2, "0.00"), 0, True, ".", , 0)
    End If
    If Me.TxtModFlg.text = "N" Then
     A_NoteID = CStr(new_id("Notes", "NoteID", "", True))
     TXT_A_NoteID.text = A_NoteID
    Else
     A_NoteID = Val(TXT_A_NoteID.text)
    End If
    
    
     rs("A_NoteID").value = Val(A_NoteID)
     
    rs.update
    
    
    
    
    
    '/////////////////////Õ”«»«  ⁄«„Â
    Dim line_no  As Integer
  

    If Me.CboPaymentType1.ListIndex = 1 Then
      Set RsNotes = New ADODB.Recordset
       RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   
        If TxtModFlg.text = "N" Then
           
            
            
           
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
 
        
        
        
    End If
    
  '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
   ' rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
   ' rs("foxy_no").value = Val(Text1.text)
'œ«∆‰ Õ”«»« 
   RsNotes.AddNew
    RsNotes("NoteID").value = A_NoteID
     RsNotes("branch_no").value = Val(Me.dcBranch.BoundText)
    RsNotes("order_no").value = TXT_order_no.text
     RsNotes("notes_all").value = Me.XPTxtID.text
    RsNotes("Note_Value").value = IIf(Not IsNumeric(XPTxtVal.text), 0, Val(XPTxtVal.text))
    'RsNotes("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
    RsNotes("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    RsNotes("too").value = IIf(TXTTo.text = "", "", Trim(TXTTo.text))
'    RsNotes("general_des").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    
    If Me.CboPayMentType.ListIndex = 0 Then
        RsNotes("BoxID").value = Val(DcboBox.BoundText)
        RsNotes("BankID").value = Null
        RsNotes("ChqueNum").value = Null
        RsNotes("DueDate").value = Null
        RsNotes("NoteCashingType").value = 0
    ElseIf Me.CboPayMentType.ListIndex = 1 Then
        RsNotes("BoxID").value = Null
        RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
        RsNotes("NoteCashingType").value = 1
        
    ElseIf Me.CboPayMentType.ListIndex = 2 Then
    RsNotes("CusID").value = DCVendor.BoundText
    
        ElseIf Me.CboPayMentType.ListIndex = 3 Then
        RsNotes("BoxID").value = Null
        RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
        RsNotes("NoteCashingType").value = 3
        
    ElseIf Me.CboPayMentType.ListIndex = 5 Then
        RsNotes("BoxID").value = Null
        RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
        RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        RsNotes("DueDate").value = Me.DtpChequeDueDate.value
        RsNotes("NoteCashingType").value = 5
        
    End If
    
    
    
    RsNotes("NoteType").value = 80
    RsNotes("NoteDate").value = XPDtbTrans.value
    RsNotes("UserID").value = user_id
    
  
 
    
    'rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
    'rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
    RsNotes("Buy").value = "0"
   ' RsNotes("Remark").value = XPMTxtRemarks.text
   RsNotes("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
    RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
    RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    RsNotes("numbering_type1").value = sand_numbering_type(8) '‰Ê⁄  —ﬁÌ„   ›« Ê—… „«·Ì…
     
    RsNotes("sanad_year").value = year(XPDtbTrans.value)
    RsNotes("sanad_month").value = Month(XPDtbTrans.value)
    RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
    RsNotes.update
    
        Dim IntDEV_Type As Integer
        Dim SngDEV_Value As Single
    line_no = 1
      LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            
            If ModAccounts.AddNewDev(LngDevID, line_no, _
                DcboCreditSide.BoundText, IIf(Not IsNumeric(XPTxtVal.text), 0, Val(XPTxtVal.text)), 1, _
                bankDes, A_NoteID, , , _
                SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , Val(Me.XPTxtID.text), , , , , , , , Val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                    
            End If
            
'„œÌ‰ Õ”«»« 
    With VSFlexGrid1
 line_no = 2
 
    For i = .FixedRows To .Rows - 1
    
Dim project_id As Integer
    
        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
       
     project_id = get_project_id(dcproject.BoundText, "expanses_account")
   
          LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            If ModAccounts.AddNewDev(LngDevID, line_no, _
                .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("Value")), 0, _
                .TextMatrix(i, .ColIndex("Des")), A_NoteID, , , _
                SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , Val(Me.XPTxtID.text), project_id, , , , , , , Val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                    
            End If
            line_no = line_no + 1
            
        End If
    Next i
End With
        
     ' TxtModFlg.text = "R"
      GoTo ll
      
      
    End If
    
    
    
    
    
    
    
'  «·«’Ê· „œÌ‰
    
   '//////////////////////////////////////Notes////////////////////////////////////
   Set RsNotes = New ADODB.Recordset
   RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
       
        Set RsDev = New ADODB.Recordset
        RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        '«·ÿ—› «·„œÌ‰
 
Dim ExpensesID As Double

 
Dim NoteID As String
  With Me.VSFlexGrid2
 

 line_no = 1
       
                'project_id = get_project_id(dcproject.BoundText, "expanses_account")
                
    For i = .FixedRows To .Rows - 1
   

        
   
        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
              '////////////////////////////////////////notes
                
               If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
               If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
               Else
               MsgBox "Cant save no value in line no:  " & i - 1, vbCritical: GoTo ErrTrap
               End If
               
               End If
                RsNotes.AddNew
                NoteID = CStr(new_id("Notes", "NoteID", "", True))
                RsNotes("NoteID").value = CStr(NoteID)
                
                 RsNotes("branch_no").value = Val(Me.dcBranch.BoundText)
                
                RsNotes("Note_Value").value = .TextMatrix(i, .ColIndex("value"))
              '  RsNotes("Remark").value = .TextMatrix(I, .ColIndex("des"))
                
                  RsNotes("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))

                RsNotes("foxy_no").value = Val(Text1.text)
                       '  If Me.CboPayMentType.ListIndex = 0 Then
                       '     rsnotes("BoxID").value = Val(DcboBox.BoundText)
                       '     rsnotes("BankID").value = Null
                       '     rsnotes("ChqueNum").value = Null
                       '     rsnotes("DueDate").value = Null
                       '     rsnotes("NoteCashingType").value = 0
                       ' ElseIf Me.CboPayMentType.ListIndex = 1 Then
                       '     rsnotes("BoxID").value = Null
                       '     rsnotes("BankID").value = Val(Me.DCboBankName.BoundText)
                       '     rsnotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                       '     rsnotes("DueDate").value = Me.DtpChequeDueDate.value
                       '     rsnotes("NoteCashingType").value = 1
                       ' End If
               
           If TXT_order_no.text <> "" Then
                RsNotes("order_no").value = TXT_order_no.text
            Else
             RsNotes("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
            End If
            
                RsNotes("CusID").value = Null
                RsNotes("NoteType").value = 80
                RsNotes("NoteDate").value = XPDtbTrans.value
                RsNotes("UserID").value = user_id
              '  RsNotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
'
               
                RsNotes("notes_all").value = Me.XPTxtID.text
                RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
                RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”·   ›« Ê—…
                RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                RsNotes("numbering_type1").value = sand_numbering_type(8) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
                
                RsNotes("sanad_year").value = year(XPDtbTrans.value)
                RsNotes("sanad_month").value = Month(XPDtbTrans.value)
                RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(.TextMatrix(I, .ColIndex("value")), "0.00"), 0, True, ".")
                
                RsNotes.update
              
              '////////////////////////////////////////notes
 
 
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            If ModAccounts.AddNewDev(LngDevID, line_no, _
                  .TextMatrix(i, .ColIndex("AccountCode")), Val(.TextMatrix(i, .ColIndex("value"))), 0, _
                 .TextMatrix(i, .ColIndex("des")), Val(NoteID), , , _
                SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , Val(.TextMatrix(i, .ColIndex("value"))), , , , , , _
                Val(Me.XPTxtID.text), , , , , Val(.TextMatrix(i, .ColIndex("id"))), Val(.TextMatrix(i, .ColIndex("groupid"))), Val(.TextMatrix(i, .ColIndex("branch_id"))), Val(Me.dcBranch.BoundText)) = False Then
                 '   GoTo ErrTrap
                    
            End If
            line_no = line_no + 1
             
        End If
    Next i
End With
    
       ' «·«’Ê· «·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
                RsNotes.AddNew
                NoteID = CStr(new_id("Notes", "NoteID", "", True))
                RsNotes("NoteID").value = CStr(NoteID)
                RsNotes("branch_no").value = Val(Me.dcBranch.BoundText)
 
                RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0)
                RsNotes("Remark").value = Me.txt_general_des
                RsNotes("foxy_no").value = Val(Text1.text)
                         If Me.CboPayMentType.ListIndex = 0 Then
                            RsNotes("BoxID").value = Val(DcboBox.BoundText)
                            RsNotes("BankID").value = Null
                            RsNotes("ChqueNum").value = Null
                            RsNotes("DueDate").value = Null
                            RsNotes("NoteCashingType").value = 0
                        ElseIf Me.CboPayMentType.ListIndex = 1 Then
                            RsNotes("BoxID").value = Null
                            RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
                            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                            RsNotes("NoteCashingType").value = 1
                      ElseIf Me.CboPayMentType.ListIndex = 2 Then
                         RsNotes("CusID").value = Val(DCVendor.BoundText)
 
                       ElseIf Me.CboPayMentType.ListIndex = 3 Then
                            RsNotes("BoxID").value = Null
                            RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
                            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                            RsNotes("NoteCashingType").value = 3
                            
                            ElseIf Me.CboPayMentType.ListIndex = 5 Then
                            RsNotes("BoxID").value = Null
                            RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
                            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                            RsNotes("NoteCashingType").value = 5
                            
                        End If
               ' RsNotes("order_no").value = txt_ORDER_NO.text
            '    RsNotes("CusID").value = Null
                RsNotes("NoteType").value = 80
                RsNotes("NoteDate").value = XPDtbTrans.value
                RsNotes("UserID").value = user_id
               ' rsnotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
                RsNotes("notes_all").value = Me.XPTxtID.text
                RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
                RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
                RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                RsNotes("numbering_type1").value = sand_numbering_type(8) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
                RsNotes("sanad_year").value = year(XPDtbTrans.value)
                RsNotes("sanad_month").value = Month(XPDtbTrans.value)
                RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
                
                RsNotes.update
                
    
                                '«·ÿ—› «·œ«∆‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
                    RsDev.AddNew
                        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
                        RsDev("DEV_ID_Line_No").value = line_no
                        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
                        RsDev("Account_Code").value = DcboCreditSide.BoundText
                        RsDev("Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
                        RsDev("Credit_Or_Debit").value = 1
                        RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text  ' .TextMatrix(I, .ColIndex("des"))
                        RsDev("RecordDate").value = Me.XPDtbTrans.value
                        RsDev("Notes_ID").value = Val(NoteID) '(XPTxtID.text)
                        RsDev("branch_id").value = Val(Me.dcBranch.BoundText)
                        RsDev("UserID").value = Me.DCboUserName.BoundText
                        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                        RsDev("notes_all").value = Me.XPTxtID.text
                      '   RsDev("project_id").value = project_id
                        
                    RsDev.update

     
     
'GoTo ll
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
 
line_no = line_no + 1
 If Me.dcproject.BoundText <> "" Then
                                '«·ÿ—› «·„œÌ‰   „’—Ê›«  «·„‘—Ê⁄
                RsNotes.AddNew
                NoteID = CStr(new_id("Notes", "NoteID", "", True))
                RsNotes("NoteID").value = CStr(NoteID)
                
          
                RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0)
                RsNotes("Remark").value = txt_general_des.text 'txtto.text
                 RsNotes("branch_no").value = Val(Me.dcBranch.BoundText)
                         If Me.CboPayMentType.ListIndex = 0 Then
                            RsNotes("BoxID").value = Val(DcboBox.BoundText)
                            RsNotes("BankID").value = Null
                            RsNotes("ChqueNum").value = Null
                            RsNotes("DueDate").value = Null
                            RsNotes("NoteCashingType").value = 0
                        ElseIf Me.CboPayMentType.ListIndex = 1 Then
                            RsNotes("BoxID").value = Null
                            RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
                            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                            RsNotes("NoteCashingType").value = 1
                                ElseIf Me.CboPayMentType.ListIndex = 2 Then
    RsNotes("CusID").value = DCVendor.BoundText
 
                          ElseIf Me.CboPayMentType.ListIndex = 3 Then
                            RsNotes("BoxID").value = Null
                            RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
                            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                            RsNotes("NoteCashingType").value = 3
                            
                            ElseIf Me.CboPayMentType.ListIndex = 5 Then
                            RsNotes("BoxID").value = Null
                            RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
                            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                            RsNotes("NoteCashingType").value = 5
                            
                        End If
               ' RsNotes("order_no").value = txt_ORDER_NO.text
            '    RsNotes("CusID").value = Null
                RsNotes("NoteType").value = 80
                RsNotes("NoteDate").value = XPDtbTrans.value
                RsNotes("UserID").value = user_id
               ' rsnotes("ExpensesID").value = .TextMatrix(I, .ColIndex("ExpensesID"))
                RsNotes("notes_all").value = Me.XPTxtID.text
                RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
                RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
                RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                RsNotes("numbering_type1").value = sand_numbering_type(8) '‰Ê⁄  —ﬁÌ„  ›« Ê—… „«·Ì…
                RsNotes("sanad_year").value = year(XPDtbTrans.value)
                RsNotes("sanad_month").value = Month(XPDtbTrans.value)
                
                RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
                
                RsNotes.update
                
                
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                Set RsDev = New ADODB.Recordset
                
        RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
                    RsDev.AddNew
                        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
                        RsDev("DEV_ID_Line_No").value = line_no
                        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
                        RsDev("Account_Code").value = dcproject.BoundText
                        RsDev("branch_id").value = Val(Me.dcBranch.BoundText)
                        RsDev("Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
                        RsDev("Credit_Or_Debit").value = 0
                        RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text ' .TextMatrix(I, .ColIndex("des"))
                        RsDev("RecordDate").value = Me.XPDtbTrans.value
                        RsDev("Notes_ID").value = Val(NoteID) '(XPTxtID.text)5
                       
                        RsDev("UserID").value = Me.DCboUserName.BoundText
                        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                        RsDev("notes_all").value = Me.XPTxtID.text
                         RsDev("project_id").value = project_id
                         
                        
                    RsDev.update
                    
 line_no = line_no + 1
  With Fg_Journal
    For i = .FixedRows To .Rows - 1
    

        
        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
              '////////////////////////////////////////notes
                
               If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
               If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
               Else
               MsgBox "Cant save enter value in line :  " & i - 1, vbCritical: GoTo ErrTrap
               End If
               
               End If
                     project_id = get_project_id(dcproject.BoundText, "expanses_account")
 
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            If ModAccounts.AddNewDev(LngDevID, line_no, _
                .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("value")), 1, _
                 .TextMatrix(i, .ColIndex("des")), Val(NoteID), , , _
                SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , setfoxy_Line, Val(Me.XPTxtID.text), project_id, , , , , , Val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                    
            End If
            line_no = line_no + 1
            
            
             
        
        
        End If
    Next i
End With
Dim sql As String
  sql = "Update notes    set note_value_by_characters='" & WriteNo(Format(Val(Me.XPTxtVal.text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & Val(TxtSerial.text)
  Cn.Execute sql
  sql = "Update   notes_all  set note_value_by_characters='" & WriteNo(Format(Val(Me.XPTxtVal.text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & Val(TxtSerial.text)
  Cn.Execute sql
 End If
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
       UpdateFixedAssetPurchaseInformations ' ÕœÌÀ »Ì«‰«  «·«’· «
       
        LblDevID.Caption = LngDevID
        lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
    End If
ll:
    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    CuurentLogdata
    Select Case Me.TxtModFlg.text
        Case "N"
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
        Else
        Msg = " Saved... " & Chr(13)
            Msg = Msg + "Do you want to enter another operation?"
        
        End If
              Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
            Cmd_Click (0)
            Exit Sub
            End If
        Case "E"
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Else
        MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        
        End If
            Fg_Journal.Enabled = False
    End Select
     'Õ›Ÿ »Ì«‰«  «·‘Ìﬂ« 
      saveChequeBoxContents1 (Val(Me.XPTxtID.text))
      
    
    
    TxtModFlg.text = "R"
End If

Exit Sub
ErrTrap:
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If
    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
    Else
       Msg = "cant save " & Chr(13)
        Msg = Msg + "Invalid entry value " & Chr(13)
        Msg = Msg + "Check data and try again"
    
    End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Else
    Msg = "Sorr.... Error during saving " & Chr(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Function UpdateFixedAssetPurchaseInformations(Optional delete As Boolean)
Dim sql As String
Dim i As Integer
Dim KhordaPrice As Double
Dim CurrentValue As Double
Dim PurcahsePrice As Double
Dim Installmentvalue As Double
  With Me.VSFlexGrid2
     For i = .FixedRows To .Rows - 1
        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
 
 
 
sql = "update FixedAssets set PurchaseDate=CONVERT(DATETIME, '" & XPDtbTrans.value & " 00:00:00', 103), PurchaseBillId=" & TxtSerial1.text & ",PurchasePrice="
           
             PurcahsePrice = Val(.TextMatrix(i, .ColIndex("value")))
              sql = sql & PurcahsePrice
           
            Dim noofinstllments As Double
              
            GetAllDataAboutFixedAsset Val(.TextMatrix(i, .ColIndex("id"))), , , , , , , , , , , , , noofinstllments, , , , , , KhordaPrice
            CurrentValue = PurcahsePrice - KhordaPrice
            sql = sql & ",CurrentValue= " & CurrentValue
            If noofinstllments = 0 Then
            noofinstllments = 0
            Else
            Installmentvalue = Round(CurrentValue / noofinstllments, 2)
            End If
            
            sql = sql & ",Installmentvalue= " & Installmentvalue
             sql = sql & ",NoteSerial=' " & Me.TxtNoteSerial.text & "'"
              sql = sql & "  where id=" & Val(.TextMatrix(i, .ColIndex("id")))
          Cn.Execute sql
          If noofinstllments <> 0 Then
          updateFixedAsseTInstallmentInformations Val(.TextMatrix(i, .ColIndex("id"))), , , , XPDtbTrans.value, , , , True, True ' ÕœÌÀ »Ì«‰«  «·«ﬁ”«ÿ
          End If
            If delete = True Then
          '  sql = "update FixedAssets NoteSerial=0,  PurchaseBillId=" & "" & ",PurchasePrice=0,Installmentvalue=0,CurrentValue=0"
            End If
            
            
        End If
        
        
    Next i
End With
End Function
Public Function save_General_cost_center(cost_center_id As String, cost_center, opr_type As String, record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
 Dim i As Integer
 Dim rs As New ADODB.Recordset
 Dim StrSQL As String
 
 StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text)
 Cn.Execute StrSQL, , adExecuteNoRecords
    
rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable

With Fg_Journal

 
.Rows = .Rows + 1
For i = .FixedRows To .Rows - 1
    If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then

        
        rs.AddNew
        rs("cost_center_id").value = cost_center_id
        rs("cost_center").value = cost_center
        rs("value").value = .TextMatrix(i, .ColIndex("value"))
        rs("depit_or_credit").value = "„œÌ‰"
        rs("opr_id").value = Me.Text1.text
        rs("kedno").value = Me.Text1.text
        rs("opr_type").value = opr_type
        rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
        rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
        rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
        rs("record_date").value = record_date
        rs.update
        
    End If
Next i
End With
rs.Close
End Function



Private Sub Undo()
On Error GoTo ErrTrap
Select Case TxtModFlg.text
    Case "N"
         clear_all Me
         Me.TxtModFlg.text = "R"
         XPBtnMove_Click (1)
    Case "E"
         rs.Find "NoteID='" & Val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst
         If rs.EOF Or rs.BOF Then
            Me.TxtModFlg.text = "R"
            Exit Sub
         End If
         Retrive
         Me.TxtModFlg.text = "R"
End Select
Exit Sub
ErrTrap:
End Sub
Private Sub Del_Trans()
Dim Msg As String
Dim StrSQL As String
On Error GoTo ErrTrap
               If SystemOptions.banks_Accounts3 = True Then
        If ChequeBoxOperations1(Val(Me.XPTxtID)) = False Then
            Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–… «·⁄„·Ì…"
            Msg = Msg & Chr(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    End If
    
      Dim noOfInstallments As Integer 'Â–« «·Ã“¡ Ì √ﬂœ „‰  ‰›Ì– «ﬁ”«ÿ «Â·«ﬂ
  Dim msgStr As String
  Dim i As Integer
  With Me.VSFlexGrid2
     For i = .FixedRows To .Rows - 1
         If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
              '////////////////////////////////////////notes
                
               noOfInstallments = CheCkInstallmentCount(Val(.TextMatrix(i, .ColIndex("id"))))
               If noOfInstallments > 0 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                        msgStr = " ·« Ì„ﬂ‰ «· ⁄œÌ·  „  ‰›Ì– «ﬁ”«ÿ ⁄·Ï «·«’·  " & Chr(13)
                        msgStr = msgStr & .TextMatrix(i, .ColIndex("AccountName")) & Chr(13)
                        msgStr = msgStr & "⁄œœ «·«ﬁ”«ÿ «·„‰›–… Õ Ï «·«‰ " & noOfInstallments
                        MsgBox msgStr, vbCritical
                        Else
                           msgStr = " Can't Edit Fixed Asset   " & Chr(13)
                        msgStr = msgStr & .TextMatrix(i, .ColIndex("AccountName")) & Chr(13)
                        msgStr = msgStr & "No Of Executed Installments " & noOfInstallments
                        MsgBox msgStr, vbCritical
                        End If
                        Exit Sub
               End If
               
               
              
        End If
        
    Next i
End With


'    UpdateFixedAssetPurchaseInformations True
    
If XPTxtID.text <> "" Then
    Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
    Msg = Msg + (TxtNoteSerial.text) & Chr(13)
    Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
     StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text)
       Cn.Execute StrSQL, , adExecuteNoRecords
 
    StrSQL = "Delete From notes Where NoteID=" & Val(TXT_A_NoteID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
            
    StrSQL = "Delete From ExpensesDetails Where NoteSerial1='" & Val(TxtSerial1.text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    
             StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & Val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
UPDATEStatusToNewAsset
       If Not rs.RecordCount < 1 Then
       CuurentLogdata ("D")
            rs.delete
            rs.MoveFirst
            If rs.RecordCount < 1 Then
                clear_all Me
                Fg_Journal.Clear flexClearScrollable, flexClearEverything
                Fg_Journal.Rows = 3
                Fg_Journal.Enabled = False
                
                VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                VSFlexGrid1.Rows = 2
                VSFlexGrid1.Enabled = False
                
                TxtModFlg_Change
                XPTxtCurrent.Caption = 0
                XPTxtCount.Caption = 0
            Else
                Retrive
            End If
        End If
    End If
Else
    clear_all Me
    Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    TxtModFlg_Change
    Exit Sub
End If
TxtModFlg_Change
Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + _
            vbExclamation, App.Title
    rs.CancelUpdate
End Sub
Function FillGridWithData()


End Function
Private Sub ReLineGrid()
Dim i As Integer
Dim IntCounter As Integer
With Fg_Journal
    For i = .FixedRows To .Rows - 1
        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
            IntCounter = IntCounter + 1
            .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            

            
        End If
    Next i
End With


IntCounter = 0
With Me.VSFlexGrid1
    For i = .FixedRows To .Rows - 1
        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
            IntCounter = IntCounter + 1
            .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
       
                    
        End If
    Next i
End With

IntCounter = 0
With Me.VSFlexGrid2
    For i = .FixedRows To .Rows - 1
        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
            IntCounter = IntCounter + 1
            .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("des")) = " ﬁÌ„… ‘—«¡ «·«’· " & .TextMatrix(i, .ColIndex("AccountName"))
                    
                    Else
                    .TextMatrix(i, .ColIndex("des")) = "PURCHASE Value Of Asset " & .TextMatrix(i, .ColIndex("AccountName"))
                    End If
                    
        End If
    Next i
End With


End Sub

Function UPDATEStatusToNewAsset()
Dim StrSQL As String
Dim i As Integer
 
With Me.VSFlexGrid2
    For i = .FixedRows To .Rows - 1
        If .TextMatrix(i, .ColIndex("id")) <> "" Then
   StrSQL = "UPDATE FixedAssets SET CurrentValue = 0,PurchaseBillId='',Installmentvalue = 0,NoteSerial='', New_or_opening=0 ,PurchasePrice=0 where  id=" & Val(.TextMatrix(i, .ColIndex("id")))
   
     Cn.Execute StrSQL
        End If
    Next i
End With


End Function
Private Sub PutData()
'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)
With Fg_Journal
    If Len(TxtDes.text) > 0 Then
        .Cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.text
        .Cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
        .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
    Else
        .Cell(flexcpData, .Row, .ColIndex("Des")) = ""
        .Cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
        .Cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
    End If
End With
End Sub

Function sand_numbering() As String
On Error Resume Next
Dim start_at As Integer
Dim end_at As Integer
Dim auto_sanad_no As String
Dim no As String
auto_sanad_no = ""
departement_name = 1
branch_no = 1
connection_string = Cn.ConnectionString
numbering.ConnectionString = connection_string
numbering.CommandType = adCmdText
numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=1"
numbering.Refresh
If numbering.Recordset.RecordCount = 0 Then
numbering_type = 0
Else
numbering_type = numbering.Recordset.Fields!numbering_id
start_at = numbering.Recordset.Fields!start_at
end_at = numbering.Recordset.Fields!end_at

End If

If numbering_type = 1 Then
detect_no.ConnectionString = connection_string
detect_no.CommandType = adCmdText
detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
detect_no.Refresh

 If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
 
 If end_at = 0 Then end_at = detect_no.Recordset.Fields!last_sand_no + 1
 
 If detect_no.Recordset.Fields!last_sand_no >= end_at Then
 sand_numbering = "error"
 Exit Function
 End If
 End If
Else
If numbering_type = 2 Then
 
detect_no.ConnectionString = connection_string
detect_no.CommandType = adCmdText
detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
detect_no.Refresh

If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
   no = Mid(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
   If end_at = 0 Then end_at = no + 1
 If no >= end_at Then
 sand_numbering = "error"
 Exit Function
 End If
 End If


Else
If numbering_type = 3 Then
 
detect_no.ConnectionString = connection_string
detect_no.CommandType = adCmdText
detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
detect_no.Refresh
If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
no = Mid(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
If end_at = 0 Then end_at = no + 1
 If no >= end_at Then
 sand_numbering = "error"
 Exit Function
 End If
 End If

 
End If
 
End If
End If

If detect_no.Recordset.RecordCount = 0 Or IsNull(detect_no.Recordset.Fields!last_sand_no) Then

                If numbering_type = 0 Then
                 ' auto_sanad_no = 1
                Else
                    If numbering_type = 1 Then
                    auto_sanad_no = start_at
                Else
                
                    If numbering_type = 2 Then
                    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & start_at

                Else
                     If numbering_type = 3 Then
                        auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & start_at

                  End If
                  End If
                  End If
                  End If

Else
                If numbering_type = 0 Then
                'auto_sanad_no = x + 1
                Else
                    If numbering_type = 1 Then
                  auto_sanad_no = detect_no.Recordset.Fields!last_sand_no + 1
                Else
                
                    If numbering_type = 2 Then
                  '  If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) Then
                   ' no = 1
                  '  auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & "1"
                  '  Else
                    no = Mid(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (no + 1)
                  '  End If
                    
                    
                      
                Else
                     If numbering_type = 3 Then
                  '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                      'no = 1
                  '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                  '    Else
                           no = Mid(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                      auto_sanad_no = Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (no + 1)

                  '    End If

                  End If
                  End If
                  End If
                  End If

End If
sand_numbering = auto_sanad_no

'MsgBox auto_sanad_no

End Function
 Function setfoxy_Line() As Double
    
Dim x As Double
x = CStr(new_id("foxy", "id1", "", True))
setfoxy_Line = x
   Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable

 
 rs("id1").value = x ' last_line_id
 
 rs.update
 
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrTrap
If KeyCode = vbKeyReturn Then
    If Me.TxtModFlg.text = "R" Then
        Cmd_Click (0)
    Else
        SendKeys "{TAB}"
    End If
End If
If Me.TxtModFlg.text = "R" Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        XPBtnMove_Click (2)
    ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        XPBtnMove_Click (1)
    ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        XPBtnMove_Click (3)
    ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        XPBtnMove_Click (0)
    End If
End If
If KeyCode = vbKeyF12 Then
    If Cmd(0).Enabled = False Then Exit Sub
    Cmd_Click (0)
End If
If KeyCode = vbKeyF11 Then
    If Cmd(1).Enabled = False Then Exit Sub
    Cmd_Click (1)
End If
If KeyCode = vbKeyF10 Then
    If Cmd(2).Enabled = False Then Exit Sub
    Cmd_Click (2)
End If
If KeyCode = vbKeyF9 Then
    If Cmd(3).Enabled = False Then Exit Sub
    Cmd_Click (3)
End If
If KeyCode = vbKeyF8 Then
    If Cmd(4).Enabled = False Then Exit Sub
    Cmd_Click (4)
End If
If Shift = 2 Then
    If KeyCode = vbKeyX Then
        If Cmd(6).Enabled = False Then Exit Sub
        Cmd_Click (6)
    End If
End If
Exit Sub
ErrTrap:
End Sub
Private Sub AddTip()
Dim Wrap As String
Dim BolRtl As Boolean
If SystemOptions.UserInterface = ArabicInterface Then
    BolRtl = True
Else
    BolRtl = False
End If
On Error GoTo ErrTrap
Wrap = Chr(13) + Chr(10)
Set TTP = New clstooltip
If BolRtl = True Then
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(0), _
        "ÃœÌœ ..." & Wrap & _
        "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(1), _
        " ⁄œÌ· ..." & Wrap & _
        "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(2), _
        "Õ›Ÿ ..." & Wrap & _
        "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & _
         "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(3), _
        " —«Ã⁄ ..." & Wrap & _
        "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & _
         "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
     With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(4), _
        "Õ–› ..." & Wrap & _
        "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(6), _
        "Œ—ÊÃ ..." & Wrap & _
        "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(1), _
        "«·√Ê· ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(0), _
        "«·”«»ﬁ ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(3), _
        "«· «·Ì ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(2), _
        "«·√ŒÌ— ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl CmdHelp, _
        "„”«⁄œ… ..." & Wrap & _
        "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & _
        "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & _
        "≈÷€ÿ Â‰«" & Wrap, True
    End With
Else
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(0), _
        "Add New Record..." & Wrap & _
        "Shortcut Key F12 OR Enter" & Wrap & _
        "OR Alt+N", BolRtl
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(1), _
        "Edit the Current Record..." & Wrap & _
        "Shortcut Key F11 " & Wrap & _
        "OR Alt+E", BolRtl
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(2), _
        "Save the New Record OR Save the Editing in the Current Record..." & Wrap & _
        "Shortcut Key F10 " & Wrap & _
        "OR Alt+S", BolRtl
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(3), _
        "Cancel the New Record OR Cancel Editing in the Current Record..." & Wrap & _
        "Shortcut Key F9 " & Wrap & _
        "OR Alt+U", BolRtl
    End With
     With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(4), _
       "Delete the Current Record..." & Wrap & _
        "Shortcut Key F8 " & Wrap & _
        "OR Alt+D", BolRtl
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl Cmd(6), _
       "Close this Screen" & Wrap & _
        "OR Alt+X", BolRtl
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(1), _
        "«·√Ê· ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(0), _
        "«·”«»ﬁ ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(3), _
        "«· «·Ì ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl XPBtnMove(2), _
        "«·√ŒÌ— ..." & Wrap & _
        "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
        " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
       .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
       .MaxWidth = 4000
       .VisibleTime = 9000
       .DelayTime = 600
       .AddControl CmdHelp, _
        "Help..." & Wrap & _
        "Display Help for this Screen" & Wrap & _
        "Shortcut Key F1" & Wrap, BolRtl
    End With
End If
Exit Sub
ErrTrap:
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim IntResult As String
Dim StrMSG As String
On Error GoTo ErrTrap
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

IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)
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

Private Sub XPCboExpensesType_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("ExpensesType", "ID", Val(Me.XPCboExpensesType.BoundText))
End If
End Sub

Private Sub XPDtbTrans_Change()
TxtSerial.text = ""
TxtSerial1.text = ""
End Sub

Private Sub XPTxtVal_Change()
 XPTxtValView.text = Format(Val(XPTxtVal.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
If SystemOptions.UserInterface = ArabicInterface Then
Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 0)

Else



'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0, , , , 1)
Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 1)

End If

    
End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
'KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
End Sub

Private Sub XPTxtVal_Validate(Cancel As Boolean)
'If Val(XPTxtVal.Text) = 0 Then
'    Set TTD = New clstooltipdemand
'    TTD.Style = TTBalloon
'    TTD.Icon = TTIconWarning
'    TTD.Centered = True
'    TTD.RightToLeft = True
'    TTD.VisibleTime = 600
'    TTD.BackColor = 0
'    TTD.Title = "ﬁÌ„… «·„’—Ê›« "
'    TTD.TipText = "»—Ã«¡ ﬂ «»… ﬁÌ„… «·„’—Ê›« "
'    TTD.PopupOnDemand = True
'    TTD.CreateToolTip XPTxtVal.hwnd
'    TTD.Show 0, XPTxtVal.Height / Screen.TwipsPerPixelX - 1    '//In Pixel only
'    Cancel = True
'Else
'    TTD.Destroy
'End If
End Sub
Private Sub ViewDataList()
Dim FrmView As FrmViewList
Dim FG As VSFlex8UCtl.vsFlexGrid
Dim StrSQL As String
Dim rs As ADODB.Recordset
Dim StrComboList As String
Dim GrdBack As ClsBackGroundPic
'Dim cProgress As ClsProgress
Dim BolFrmLoaded As Boolean
Set FrmView = New FrmViewList
Set FG = FrmView.vsfGroup1.vsFlexGrid

With FG
    .Cols = 18
    .RowHeightMin = 320
    .ExplorerBar = flexExSortShowAndMove
    .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
    .ColKey(0) = "NoteID"
    .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
    .ColKey(1) = "NoteSerial"
    .TextMatrix(0, 2) = "«· «—ÌŒ"
    .ColKey(2) = "NoteDate"
    .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
    .ColKey(3) = "Name"
    .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
    .ColKey(4) = "Note_Value"
    .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
    .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
    .ColKey(5) = "BoxName"
    .TextMatrix(0, 6) = "„·«ÕŸ« "
    .ColKey(6) = "Remark"
    .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
    .ColKey(7) = "UserName"
    
    StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & _
    "Remark, UserName From ExpensesReport"
    StrSQL = StrSQL + " Order By NoteID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    'Â‰« Ìﬂ » ﬂÊœ ·⁄„· „⁄œ·  Õ„Ì· «·»Ì«‰« 
    '------------------------------------
    '
    '
    '
    '
    
    '------------------------------------
    Set .DataSource = rs
    .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
    .ColKey(0) = "NoteID"
    .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
    .ColKey(1) = "NoteSerial"
    .TextMatrix(0, 2) = "«· «—ÌŒ"
    .ColKey(2) = "NoteDate"
    .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
    .ColKey(3) = "Name"
    .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
    .ColKey(4) = "Note_Value"
    .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
    .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
    .ColKey(5) = "BoxName"
    .TextMatrix(0, 6) = "„·«ÕŸ« "
    .ColKey(6) = "Remark"
    .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
    .ColKey(7) = "UserName"
    
    'Rs.Close
    'Set Rs = Nothing
    .AutoSize 0, .Cols - 1, False
End With
Set GrdBack = New ClsBackGroundPic
FrmView.vsfGroup1.vsFlexGrid.WallPaper = GrdBack.Picture
FrmView.vsfGroup1.SetRTL = True
FrmView.vsfGroup1.TotalOnColKey = "Note_Value"
FrmView.vsfGroup1.sql = StrSQL
FrmView.vsfGroup1.ShowTreeGroups = True
FrmView.vsfGroup1.update
FrmView.SetDblClickRetrun Me, "NoteID"
FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·»Ì«‰«  «·„’—Ê›« "
FrmView.Show
End Sub
Private Sub ChangeLang()
Dim XPic As IPictureDisp
Set XPic = Me.XPBtnMove(1).ButtonImage
Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
Set Me.XPBtnMove(2).ButtonImage = XPic
Set XPic = Me.XPBtnMove(0).ButtonImage
Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
Set Me.XPBtnMove(3).ButtonImage = XPic
'LblValue.Visible = False
 
lbl(24).Caption = "Hint."
lbl(25).Caption = "This Window Allow Purchase Of Fixed Assets"

lbl(23).Caption = "Invoice Type"
Label3.Caption = "GL No."
lbl(14).Caption = "Project#"
'Label1.Caption = "Manual #"
Me.ALLButton1.Caption = "Cost Center"
lbl(15).Caption = "Payment Method"
lbl(16).Caption = "Box Name"
lbl(20).Caption = "General Des"
lbl(21).Caption = "Order No:"
Label1.Caption = "Branch"
lbl(26).Caption = "Account"

Label8.Caption = "General C. C."

With Me.CboPayMentType
    .Clear
    .AddItem "Cash"
    .AddItem "Cheque"
    .AddItem "Credit"
    .AddItem "Transfer"
    .AddItem "Account"
    .AddItem "P Cheque"
End With

With Me.CboPaymentType1
    .Clear
    .AddItem "Expenses"
    .AddItem "Accounts"
    .AddItem "Fixed Asset Purchase"
End With
CmdRemove.Caption = "Delete Row"
Me.Caption = "Fixed Asset Purchase Invoice"
Me.Ele.Caption = Me.Caption


Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
Me.lbl(4).Caption = "Operation ID"
Me.lbl(1).Caption = "Operation Date"
Me.lbl(3).Caption = "Expenses Type"
Me.lbl(2).Caption = "Total"
Me.lbl(0).Caption = "Vendor Bill#"
Me.lbl(5).Caption = "Remarks"
Me.lbl(8).Caption = "Issued By."
Me.lbl(7).Caption = "Current Record."

Fra.Caption = "GL"
lbl(11).Caption = "GL#"
lbl(13).Caption = "interval"
lbl(9).Caption = "Depit"
lbl(10).Caption = "Credit"
lbl(17).Caption = "Bank"
lbl(18).Caption = "Cheque#"
lbl(19).Caption = "Due Date"
lbl(22).Caption = "Vendor"





Me.Cmd(0).Caption = "&New"
Me.Cmd(1).Caption = "&Edit"
Me.Cmd(2).Caption = "&Save"
Me.Cmd(3).Caption = "&Undo"
Me.Cmd(4).Caption = "&Delete"
Me.Cmd(5).Caption = "Sear&ch"
Me.Cmd(6).Caption = "E&xit"
Me.Cmd(7).Caption = "&Table View"
Cmd(8).Caption = "Print"
Cmd(9).Caption = "Cheque Print"
Cmd(10).Caption = "GL Print "

Me.CmdHelp.Caption = "&Help"

With Me.Fg_Journal
.TextMatrix(0, .ColIndex("LineNo")) = "Index"
.TextMatrix(0, .ColIndex("AccountName")) = " Expenses Name"
.TextMatrix(0, .ColIndex("value")) = "value"
.TextMatrix(0, .ColIndex("des")) = "description"
.TextMatrix(0, .ColIndex("opr_fullcode")) = "Operation"

End With


With VSFlexGrid1
.TextMatrix(0, .ColIndex("LineNo")) = "Index"
.TextMatrix(0, .ColIndex("AccountName")) = " Account Name"
.TextMatrix(0, .ColIndex("Account_Serial")) = " Account Code  "
.TextMatrix(0, .ColIndex("Value")) = "value"
 
End With

With VSFlexGrid2
.TextMatrix(0, .ColIndex("LineNo")) = "Index"
.TextMatrix(0, .ColIndex("AccountName")) = " Fixed Asset Name"

.TextMatrix(0, .ColIndex("Value")) = "value"
.TextMatrix(0, .ColIndex("Des")) = "  Des.  "
End With

End Sub
