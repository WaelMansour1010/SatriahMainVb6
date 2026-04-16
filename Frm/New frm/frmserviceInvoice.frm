VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form frmserviceInvoice 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›« Ê—… Œœ„Ì…"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17010
   HelpContextID   =   280
   Icon            =   "frmserviceInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   17010
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox txtTotalFines 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   122
      Top             =   7080
      Width           =   945
   End
   Begin VB.CommandButton cmdCreateTransa 
      Caption         =   "«‰‘«¡ ‰”Œ ÿ»ﬁ« ··‰”Œ… «·Õ«·Ì…"
      Height          =   855
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   120
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox txtRepeatNo 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   12720
      RightToLeft     =   -1  'True
      TabIndex        =   118
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox txtVatTotal 
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
      Height          =   360
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   117
      Top             =   7020
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   14760
      ScaleHeight     =   195
      ScaleWidth      =   735
      TabIndex        =   116
      Top             =   8760
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox XPTxtVal2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2340
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   106
      Top             =   6540
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.TextBox XPTxtValView2 
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
      Height          =   360
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   105
      Top             =   6510
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   960
      Width           =   16935
      Begin VB.OptionButton ComResid 
         Alignment       =   1  'Right Justify
         Caption         =   "€Ì— Œ«÷⁄"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   12060
         RightToLeft     =   -1  'True
         TabIndex        =   155
         Top             =   3450
         Width           =   1095
      End
      Begin VB.OptionButton ComResid 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ«÷⁄"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   11100
         RightToLeft     =   -1  'True
         TabIndex        =   154
         Top             =   3450
         Width           =   975
      End
      Begin VB.TextBox TXTNewNO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   14010
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   152
         Top             =   3360
         Width           =   1230
      End
      Begin VB.TextBox txtVendor 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   15360
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   150
         Top             =   3360
         Width           =   1230
      End
      Begin VB.TextBox txt_Currency_rate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -240
         RightToLeft     =   -1  'True
         TabIndex        =   146
         Text            =   "1"
         Top             =   1665
         Width           =   765
      End
      Begin VB.TextBox TXTIban 
         Height          =   495
         Left            =   -1440
         TabIndex        =   145
         Top             =   2940
         Width           =   3375
      End
      Begin VB.TextBox txtFile 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   139
         Top             =   0
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VSFlex8LCtl.VSFlexGrid grdFiles 
         Height          =   435
         Left            =   10800
         TabIndex        =   138
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
         _cx             =   3413
         _cy             =   767
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmserviceInvoice.frx":038A
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
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
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
      Begin VB.TextBox txtMainPath 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10800
         RightToLeft     =   -1  'True
         TabIndex        =   137
         Text            =   "Text8"
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdLoadFile 
         Caption         =   " Õ„Ì· «·„·›..."
         Height          =   375
         Left            =   11040
         TabIndex        =   135
         Top             =   480
         Width           =   1515
      End
      Begin VB.CommandButton cmdSelectFile 
         Caption         =   " ÕœÌœ «·„·›..."
         Height          =   345
         Left            =   11040
         RightToLeft     =   -1  'True
         TabIndex        =   133
         Top             =   120
         Width           =   1500
      End
      Begin VB.TextBox txtContractNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   130
         Top             =   1230
         Width           =   1455
      End
      Begin VB.TextBox txtRequestNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   128
         Top             =   630
         Width           =   1455
      End
      Begin VB.CheckBox AkarPayCheck 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ’—›«  ⁄ﬁ«—Ì…"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   14880
         RightToLeft     =   -1  'True
         TabIndex        =   115
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtManulaVat 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   11400
         TabIndex        =   113
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox TxtNoteserial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox TxtOrderID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -120
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   30
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ComboBox CBoBasedON 
         Height          =   315
         Left            =   2040
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1920
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   30
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Text            =   "Text1"
         Top             =   -690
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox CboPaymentType 
         Height          =   315
         Left            =   6120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
      Begin VB.Frame FraNote 
         BackColor       =   &H00E2E9E9&
         Height          =   2565
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   1080
         Width           =   4635
         Begin VB.TextBox TxtChequeNumber 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   1320
            Width           =   3285
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   240
            Width           =   705
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   600
            Width           =   705
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   960
            Width           =   705
         End
         Begin MSComCtl2.DTPicker DtpChequeDueDate 
            Height          =   315
            Left            =   30
            TabIndex        =   10
            Top             =   1740
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Format          =   113180673
            CurrentDate     =   39614
         End
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   960
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   0
            TabIndex        =   7
            Top             =   600
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCVendor 
            Height          =   315
            Left            =   0
            TabIndex        =   6
            Top             =   240
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCar 
            Height          =   315
            Left            =   30
            TabIndex        =   102
            Top             =   2160
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„⁄œ…"
            Height          =   285
            Index           =   28
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·Œ“Ì‰…"
            Height          =   285
            Index           =   16
            Left            =   3270
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·»‰ﬂ"
            Height          =   285
            Index           =   17
            Left            =   3270
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   990
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   285
            Index           =   18
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
            Height          =   285
            Index           =   19
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   285
            Index           =   22
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2040
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1590
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox TxtSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox txt_general_des 
         Alignment       =   1  'Right Justify
         Height          =   765
         Left            =   0
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   2580
         Width           =   4755
      End
      Begin VB.TextBox txt_ORDER_NO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2310
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox CboPaymentType1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7920
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   510
         Width           =   1455
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   -240
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   -90
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text7 
         DataField       =   "id"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   960
         TabIndex        =   66
         Text            =   "Text2"
         Top             =   -570
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TXT_A_NoteID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Text            =   "Text8"
         Top             =   3150
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   5940
         TabIndex        =   0
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   112001025
         CurrentDate     =   38784
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   7
         Left            =   -210
         TabIndex        =   81
         Top             =   3390
         Visible         =   0   'False
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
         Left            =   2040
         TabIndex        =   5
         Top             =   1230
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcCostCenter 
         Bindings        =   "frmserviceInvoice.frx":043D
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   870
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
         Bindings        =   "frmserviceInvoice.frx":0452
         Height          =   315
         Left            =   2040
         TabIndex        =   1
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
      Begin ALLButtonS.ALLButton CmdDes 
         Height          =   345
         Left            =   2640
         TabIndex        =   110
         Tag             =   "Delete Row"
         Top             =   3360
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "‰”Œ «·‘—Õ ·ﬂ· «·”ÿÊ—"
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
         BCOL            =   16777088
         BCOLO           =   16777088
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmserviceInvoice.frx":0467
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker txtFromDate 
         Height          =   315
         Left            =   15435
         TabIndex        =   123
         Top             =   2280
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   112001025
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker txtToDate 
         Height          =   315
         Left            =   12210
         TabIndex        =   124
         Top             =   2265
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Format          =   112001025
         CurrentDate     =   38784
      End
      Begin Dynamic_Byte.NourHijriCal txtToDateH 
         Height          =   315
         Left            =   10800
         TabIndex        =   125
         Top             =   2265
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal txtFromDateH 
         Height          =   315
         Left            =   13920
         TabIndex        =   132
         Top             =   2280
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo DcCurrency 
         Height          =   315
         Left            =   600
         TabIndex        =   147
         Top             =   1650
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker txtDateRec 
         Height          =   360
         Left            =   0
         TabIndex        =   149
         Top             =   2040
         Visible         =   0   'False
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   635
         _Version        =   393216
         Format          =   114491393
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·⁄ﬁœ «·„ÊÕœ"
         Height          =   315
         Index           =   34
         Left            =   13680
         RightToLeft     =   -1  'True
         TabIndex        =   153
         Top             =   3090
         Width           =   1530
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·›‰œ— «·„ﬁ«Ê·"
         Height          =   315
         Index           =   33
         Left            =   15030
         RightToLeft     =   -1  'True
         TabIndex        =   151
         Top             =   3090
         Width           =   1530
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "««·⁄„·…"
         Height          =   300
         Index           =   65
         Left            =   930
         RightToLeft     =   -1  'True
         TabIndex        =   148
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Caption         =   "Estimated Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   10920
         TabIndex        =   140
         Top             =   1620
         Visible         =   0   'False
         Width           =   5745
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  —ﬁ„ «·⁄ﬁœ"
         Height          =   255
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   131
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·„ÿ«·»…"
         Height          =   255
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   129
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„‰"
         Height          =   255
         Index           =   31
         Left            =   14940
         RightToLeft     =   -1  'True
         TabIndex        =   127
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï"
         Height          =   375
         Index           =   76
         Left            =   11280
         RightToLeft     =   -1  'True
         TabIndex        =   126
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«œŒ«· «·‰”»… «·ÌœÊÌ…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   148
         Left            =   12600
         TabIndex        =   114
         Top             =   2760
         Width           =   1800
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  «·—ﬁ„ «·ÌœÊÌ "
         Height          =   255
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰«¡ ⁄·Ï"
         Height          =   195
         Index           =   26
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   1950
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·›« Ê—…"
         Height          =   285
         Index           =   4
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„’—Ê›« "
         Height          =   285
         Index           =   3
         Left            =   18960
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   1110
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   135
         Width           =   555
      End
      Begin VB.Image ImgNote 
         Height          =   240
         Left            =   -240
         Picture         =   "frmserviceInvoice.frx":0483
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
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   1230
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
         Height          =   195
         Index           =   15
         Left            =   9420
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ ›« Ê—… «·„Ê—œ"
         Height          =   285
         Index           =   0
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   1590
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
         Height          =   255
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·‘—Õ «·⁄«„"
         Height          =   285
         Index           =   20
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   2790
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   405
         Index           =   21
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   2190
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·›« Ê—…"
         Height          =   285
         Index           =   23
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   510
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1125
         Left            =   12840
         Top             =   360
         Width           =   2415
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
         Left            =   13440
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "«·›« Ê—… «·Œœ„Ì… ÊÂÌ  Œ’ ﬂ· «·„»Ì⁄«  «·‰ﬁœÌ… «Ê «·«Ã·… Ê«· Ï ·Ì” ·Â« «’‰«› „ÕœœÂ Ê·« Ì‰ Ã ⁄‰Â«  √ÀÌ— „Œ“‰Ì"
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
         Height          =   1080
         Index           =   25
         Left            =   13200
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   360
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
      TabIndex        =   60
      Top             =   240
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkLastAccount 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   195
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   59
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
      TabIndex        =   58
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   7890
      Width           =   1905
   End
   Begin VB.TextBox Txt_Numorder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   18
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
      TabIndex        =   42
      Top             =   9420
      Width           =   6465
      Begin MSDataListLib.DataCombo DcboCreditSide 
         Height          =   315
         Left            =   90
         TabIndex        =   46
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   44
         Top             =   270
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
         TabIndex        =   50
         Top             =   570
         Width           =   1485
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   45
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
         TabIndex        =   43
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   25
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   645
      Left            =   18840
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
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
      TabIndex        =   17
      Top             =   6930
      Width           =   1305
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   855
      Left            =   0
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   30
      Width           =   17055
      _cx             =   30083
      _cy             =   1508
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
      Picture         =   "frmserviceInvoice.frx":0A0D
      Caption         =   "  ›« Ê—… Œœ„Ì…    "
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
      Begin VB.ComboBox DefaultInvoicetype 
         Height          =   315
         ItemData        =   "frmserviceInvoice.frx":16E7
         Left            =   12360
         List            =   "frmserviceInvoice.frx":16E9
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   142
         Top             =   360
         Width           =   1890
      End
      Begin VB.CheckBox ChkPurchaseFixedAssets 
         Alignment       =   1  'Right Justify
         Caption         =   "›« Ê—… ‘—«¡ «’·"
         Height          =   195
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   98
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox oldTxtSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1695
         TabIndex        =   21
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
         ButtonImage     =   "frmserviceInvoice.frx":16EB
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
         TabIndex        =   22
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
         ButtonImage     =   "frmserviceInvoice.frx":1A85
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
         TabIndex        =   23
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
         ButtonImage     =   "frmserviceInvoice.frx":1E1F
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
         TabIndex        =   24
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
         ButtonImage     =   "frmserviceInvoice.frx":21B9
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
         Top             =   -30
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
      Begin MSDataListLib.DataCombo DCDocTypes 
         Height          =   315
         Left            =   9420
         TabIndex        =   143
         Top             =   420
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·›« Ê—…"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   122
         Left            =   10590
         TabIndex        =   144
         Top             =   450
         Width           =   1095
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   6600
         Picture         =   "frmserviceInvoice.frx":2553
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
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
         TabIndex        =   41
         Top             =   510
         Width           =   5445
      End
   End
   Begin MSDataListLib.DataCombo XPCboExpensesType 
      Height          =   315
      Left            =   18240
      TabIndex        =   16
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
      Left            =   8280
      TabIndex        =   27
      Top             =   7890
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   10140
      TabIndex        =   33
      Top             =   7380
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      Height          =   375
      Index           =   1
      Left            =   9330
      TabIndex        =   34
      Top             =   7380
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      Height          =   375
      Index           =   2
      Left            =   8520
      TabIndex        =   35
      Top             =   7380
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      Height          =   375
      Index           =   3
      Left            =   7725
      TabIndex        =   36
      Top             =   7380
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      Height          =   375
      Index           =   4
      Left            =   6900
      TabIndex        =   37
      Top             =   7380
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      Height          =   375
      Index           =   6
      Left            =   30
      TabIndex        =   38
      Top             =   7380
      Width           =   735
      _ExtentX        =   1296
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
      Height          =   375
      Left            =   870
      TabIndex        =   39
      Top             =   7380
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      Height          =   375
      Index           =   5
      Left            =   6060
      TabIndex        =   40
      Top             =   7380
      Width           =   735
      _ExtentX        =   1296
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   255
      Left            =   9930
      TabIndex        =   51
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmserviceInvoice.frx":61BB
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
      Left            =   5220
      TabIndex        =   54
      Top             =   7320
      Width           =   735
      _ExtentX        =   1296
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
      TabIndex        =   55
      Top             =   8880
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
      Height          =   285
      Left            =   10050
      TabIndex        =   56
      Tag             =   "Delete Row"
      Top             =   6750
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
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
      MICON           =   "frmserviceInvoice.frx":61D7
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
      Left            =   4200
      TabIndex        =   61
      Top             =   7800
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
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   3120
      TabIndex        =   99
      Top             =   7800
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdPrintForms 
      CausesValidation=   0   'False
      Height          =   465
      Index           =   0
      Left            =   4350
      TabIndex        =   103
      Top             =   7290
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   820
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "›« Ê—… €Ì— „‰ Ÿ„"
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
   Begin ImpulseButton.ISButton CmdPrintForms 
      CausesValidation=   0   'False
      Height          =   465
      Index           =   2
      Left            =   2610
      TabIndex        =   107
      Top             =   7290
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   820
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "›« Ê—… „‰ Ÿ„"
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
   Begin ImpulseButton.ISButton CmdPrintForms 
      CausesValidation=   0   'False
      Height          =   465
      Index           =   1
      Left            =   3480
      TabIndex        =   108
      Top             =   7290
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   820
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "›« Ê—… »Ì‰ «·„” Êœ⁄« "
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
   Begin ImpulseButton.ISButton CmdPrintForms 
      CausesValidation=   0   'False
      Height          =   465
      Index           =   3
      Left            =   1680
      TabIndex        =   109
      Top             =   7290
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   820
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "›« Ê—… «ÌÃ«— ÌÊ„Ì"
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
   Begin VSFlex8Ctl.VSFlexGrid grdExcel 
      Height          =   510
      Index           =   0
      Left            =   15390
      TabIndex        =   134
      Top             =   8190
      Visible         =   0   'False
      Width           =   3885
      _cx             =   6853
      _cy             =   900
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
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   5
      Rows            =   2
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmserviceInvoice.frx":61F3
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
      ExplorerBar     =   3
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
   Begin VSFlex8Ctl.VSFlexGrid tmpGrd 
      Height          =   615
      Left            =   12930
      TabIndex        =   136
      Top             =   8430
      Visible         =   0   'False
      Width           =   1380
      _cx             =   2434
      _cy             =   1085
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
      BackColor       =   8421631
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   8421631
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
      Rows            =   50
      Cols            =   40
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   2100
      Left            =   60
      TabIndex        =   156
      Top             =   4770
      Width           =   16755
      _cx             =   29554
      _cy             =   3704
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   128
      FrontTabColor   =   14871017
      BackTabColor    =   8454143
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "Œœ„« |„’—Ê›« "
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin VB.Frame Frame2 
         Height          =   1680
         Left            =   17400
         RightToLeft     =   -1  'True
         TabIndex        =   163
         Top             =   45
         Width           =   16665
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   1380
            Left            =   150
            TabIndex        =   164
            Top             =   150
            Visible         =   0   'False
            Width           =   16395
            _cx             =   28919
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   18
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmserviceInvoice.frx":6493
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
            Begin VB.Frame Frame3 
               Caption         =   "Õœœ —ﬁ„ «·ﬁÌœ «·„—«œ ‰”Œ…"
               Height          =   1215
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   3720
               Visible         =   0   'False
               Width           =   4215
               Begin VB.CommandButton Command5 
                  Caption         =   "‰”Œ"
                  Height          =   255
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   178
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   177
                  Top             =   240
                  Width           =   2175
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  Height          =   255
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   179
                  Top             =   240
                  Width           =   1335
               End
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   3915
               Left            =   2550
               RightToLeft     =   -1  'True
               ScaleHeight     =   3915
               ScaleWidth      =   9405
               TabIndex        =   165
               Top             =   810
               Visible         =   0   'False
               Width           =   9405
               Begin VB.TextBox TxtDese 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   1485
                  Left            =   120
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   169
                  Top             =   2040
                  Width           =   8955
               End
               Begin VB.TextBox txtcodesub 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   5400
                  RightToLeft     =   -1  'True
                  TabIndex        =   168
                  Top             =   3600
                  Width           =   855
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "Add des"
                  Height          =   255
                  Left            =   7440
                  RightToLeft     =   -1  'True
                  TabIndex        =   167
                  Top             =   3600
                  Width           =   1350
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "Call des"
                  Height          =   255
                  Left            =   6240
                  RightToLeft     =   -1  'True
                  TabIndex        =   166
                  Top             =   3600
                  Width           =   1095
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   3900
                  Left            =   120
                  TabIndex        =   170
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
                     TabIndex        =   171
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
                     TabIndex        =   172
                     Top             =   0
                     Width           =   2445
                  End
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Code"
                  Height          =   495
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   175
                  Top             =   3480
                  Width           =   735
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  Height          =   495
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   174
                  Top             =   1200
                  Width           =   975
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Code"
                  Height          =   255
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   173
                  Top             =   1320
                  Width           =   735
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Eleqq 
         Height          =   1680
         Index           =   2
         Left            =   45
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   45
         Width           =   16665
         _cx             =   29395
         _cy             =   2963
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
            Height          =   360
            Left            =   1590
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   182
            Top             =   1290
            Width           =   1425
         End
         Begin VB.TextBox txtTotal 
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
            Height          =   360
            Left            =   3870
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   181
            Top             =   1260
            Width           =   1425
         End
         Begin VB.TextBox txtPayAmount 
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
            Height          =   360
            Left            =   0
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   180
            Top             =   1290
            Width           =   1425
         End
         Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
            Height          =   1260
            Left            =   0
            TabIndex        =   158
            Top             =   0
            Width           =   16800
            _cx             =   29633
            _cy             =   2222
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   33
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmserviceInvoice.frx":676F
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
            Begin VB.PictureBox PicDes 
               BorderStyle     =   0  'None
               Height          =   1635
               Left            =   2880
               RightToLeft     =   -1  'True
               ScaleHeight     =   1635
               ScaleWidth      =   2925
               TabIndex        =   159
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
                  TabIndex        =   160
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
                  TabIndex        =   161
                  Top             =   0
                  Width           =   2445
               End
            End
            Begin VDSCOMBOLibCtl.SmartCombo CboDes 
               Height          =   315
               Left            =   4200
               TabIndex        =   162
               ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
               Top             =   840
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
               Picture         =   "frmserviceInvoice.frx":6C74
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   285
            Index           =   2
            Left            =   2910
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   285
            Index           =   30
            Left            =   5430
            RightToLeft     =   -1  'True
            TabIndex        =   183
            Top             =   1380
            Width           =   795
         End
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„”œœ"
      Height          =   285
      Index           =   32
      Left            =   1650
      RightToLeft     =   -1  'True
      TabIndex        =   141
      Top             =   6390
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·€—«„« "
      Height          =   285
      Index           =   29
      Left            =   11160
      RightToLeft     =   -1  'True
      TabIndex        =   121
      Top             =   6840
      Width           =   555
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄œœ «·‰”Œ"
      Height          =   375
      Left            =   13560
      RightToLeft     =   -1  'True
      TabIndex        =   119
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "«·—’Ìœ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   15120
      RightToLeft     =   -1  'True
      TabIndex        =   112
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label LblLink 
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
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   12000
      RightToLeft     =   -1  'True
      TabIndex        =   111
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Label lblValue2 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Height          =   405
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   104
      Top             =   6270
      Visible         =   0   'False
      Width           =   5835
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
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
      Index           =   27
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   97
      Top             =   8280
      Width           =   7155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   390
      Index           =   8
      Left            =   9945
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   7905
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   255
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Height          =   405
      Left            =   5670
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   6900
      Width           =   4095
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   7890
      Width           =   555
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   7890
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
      TabIndex        =   29
      Top             =   7890
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
      TabIndex        =   28
      Top             =   7890
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·«„—"
      Height          =   285
      Index           =   5
      Left            =   11400
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   2520
      Width           =   1515
   End
End
Attribute VB_Name = "frmserviceInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A410921053-658C-192080>>Smart Combo
Option Explicit
Dim startTime As Date
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim numbering_type As Integer
Dim departement_name  As String
Dim branch_no  As String
Dim RsNotes As ADODB.Recordset
Dim BolEditOnMainAccounts As Boolean
 Dim Dcombos As ClsDataCombos
Dim isCmdCreateTransa As Boolean
Dim s As String
Dim zatcaStatus As Integer
Dim mIndexVat As Integer
Dim Export As Integer
Public mTypeInvoice As Integer

'***************
Dim allExcelRows As Long
Dim AllFinshedRows As Long
 '**********
Dim mIsSaveFromExcel As Boolean
 Private FSO    As New Scripting.FileSystemObject
 
 Private FirstTime As Date
  Private FirstTimef As Boolean

  Public Function StringDotFormat(ByVal StrFormat As String, _
                                ParamArray aryPlaceHolders()) As String

    Dim intPlaceHolderIndex As Integer
    Dim strOutput As String
    strOutput = StrFormat
    For intPlaceHolderIndex = LBound(aryPlaceHolders) To UBound(aryPlaceHolders)
        strOutput = Replace(strOutput, "{" & intPlaceHolderIndex & "}", aryPlaceHolders(intPlaceHolderIndex) & "")
    Next
    StringDotFormat = strOutput
End Function
'Function LogMessage(ByRef Message As String)
'    '<CsTrace>
'    Dim sCodeSMARTParams As String
'    sCodeSMARTParams = AxCsDumpParamValue("Message", Message)
'
'    AxCsTrace "Dynamic_Byte", "frmserviceInvoice", "LogMessage (" & sCodeSMARTParams & ")", ProcEnter
'    '</CsTrace>
'    'Dim FSO As Object
'    Dim LogFileName As String
'    Dim File        As Object
'
'    If FirstTimef Then
'    Dim Diff As Long
'
'        LogFileName = "C:\log\MyLog.txt"
'        Set File = FSO.OpenTextFile(LogFileName, ForAppending, False, TristateTrue)
'        File.WriteLine Format(Now, "hh:mm:ss") & ": " & Message
'        File.Close
'    End If
'    FirstTime = Now
'    FirstTimef = True
'
'    '<CsTrace>
'    AxCsTrace "Dynamic_Byte", "frmserviceInvoice", "LogMessage", ProcExit
'    '</CsTrace>
'End Function
Public Function GetTimeHour(ByVal inSec As Double) As String
    Dim ss As Boolean
      Dim aHr As Double
     Dim aMin As Double
        Dim aSec As Double
    ss = (inSec < 0)
    If (inSec <> 0) Then
        inSec = IIf(inSec < 0, -1, 1) * inSec
        aHr = Fix(inSec / 3600)
        aMin = Fix((inSec - (aHr * 3600)) / 60)
        aSec = inSec - (aHr * 3600) - (aMin * 60)
        GetTimeHour = IIf(ss, "-", "") & Format(aHr, "00000") & ":" & Format(aMin, "00") & ":" & Format(aSec, "00")
    Else
        GetTimeHour = "00000:00:00"
    End If
End Function
Private Sub cmdCreateTransa_Click()

    Dim i As Long
            
            
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            For i = 1 To val(txtRepeatNo)
            TxtModFlg.text = "N"
            isCmdCreateTransa = True
            
            
               
    
                 XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
               '  Me.TxtNoteSerial.Text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
                Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=85"))
                Me.DCboUserName.BoundText = user_id
                '        XPDtbTrans.SetFocus
              
             
              
            
                setfoxy
                TxtSerial1.text = ""
                SaveData "", ""
        Next
        isCmdCreateTransa = False
        MsgBox " „ «‰‘«¡ ⁄œœ " & txtRepeatNo & "  Õ—ﬂ«  "
End Sub

Private Sub CmdDes_Click()
Dim i As Integer
 With Me.Fg_Journal
        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
             
              .TextMatrix(i, .ColIndex("Des")) = txt_general_des.text
             .cell(flexcpData, i, .ColIndex("Des")) = txt_general_des.text
 
            



            End If

        Next i

    End With
    
End Sub
Function saveChequeBoxContents1(NoteID As Double)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords

    If SystemOptions.banks_Accounts3 = False Then
        Exit Function
    End If
 
    'rs.Open "TblChecqueBoxContent1", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  StrSQL = "SELECT     * from dbo.TblChecqueBoxContent1 Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
    If CboPayMentType.ListIndex = 1 Then
        rs.AddNew
        rs("noteid").value = NoteID
     
        rs("RecordDate").value = XPDtbTrans.value
        rs("DueDate").value = DtpChequeDueDate.value
        rs("BankID").value = val(DcboBankName.BoundText)
        rs("BankName").value = DcboBankName.text
        
        rs("ChequeNo").value = TxtChequeNumber.text
        rs("ChequeValue").value = val(XPTxtVal.text)
    
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

    If Not IsNumeric(Text1.text) Then
        Exit Sub
    End If
    'If Me.TxtModFlg.text = "N" Then
    opr_id = val(Me.Text1.text)
    'Else
    'opr_id = TxtDEV_NO.text
    'End If
If CboPaymentType1.ListIndex = 0 Then
    If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) = "" Then
        If Not val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE"))) = 0 Then

            marakes_taklefa_tawze3.show
            
            marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
            marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
            marakes_taklefa_tawze3.lineno = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        
        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… ", vbCritical
            Else
                MsgBox "Enter Value First ", vbCritical
            End If

            Exit Sub
        End If
            
    End If
Else
    If Not VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("AccountCode")) = "" Then
        If Not val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("VALUE"))) = 0 Then

            marakes_taklefa_tawze3.show
            
            marakes_taklefa_tawze3.value.Caption = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("VALUE")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("AccountCode"))
            marakes_taklefa_tawze3.account_name = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("AccountName"))
            marakes_taklefa_tawze3.lineno = VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("LineNo1"))
        
        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… ", vbCritical
            Else
                MsgBox "Enter Value First ", vbCritical
            End If

            Exit Sub
        End If
            
    End If
End If
    marakes_taklefa_tawze3.opr_type = "›« Ê—… „«·Ì…"
    marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
    marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
    marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
     If CboPaymentType1.ListIndex = 0 Then
     marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
    Else
     marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & VSFlexGrid1.TextMatrix(Fg_Journal.Row, VSFlexGrid1.ColIndex("AccountCode")) & "' and  line_no=" & VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("LineNo1"))
    End If
    
    marakes_taklefa_tawze3.Adodc3.Refresh
    '    Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"

    Exit Sub
ErrTrap:
End Sub

Private Sub CBoBasedON_Click()
    CBoBasedON_Change
End Sub

Private Sub CboDes_AfterAutoCloseUp()
    PutData
    CboDes.Visible = False
End Sub

Private Sub CboPayMentType_Change()
    Dim s As String
    Dim rsDummy As ADODB.Recordset
    If Me.TxtModFlg.text = "E" Then
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        Me.DcboBox.text = ""
        DCVendor.text = ""
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
    ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 3 Then
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
     
    ElseIf Me.CboPayMentType.ListIndex = 2 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = False
    
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DcboBox.Enabled = False
        Me.DCVendor.Enabled = True
            Me.DtpChequeDueDate.Enabled = True


    ElseIf Me.CboPayMentType.ListIndex = 4 Then
        
        s = "SELECT BankId FROM TblPaymentType WHERE PaymentName LIKE '%" & "›Ì“«" & "%'"
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
        If Not rsDummy.EOF Then
            DcboBankName.BoundText = val(rsDummy!BankID & "")
        End If
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = False
    
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DcboBox.Enabled = False
        Me.DCVendor.Enabled = False
            Me.DtpChequeDueDate.Enabled = True
            
    ElseIf Me.CboPayMentType.ListIndex = 5 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = False
        s = "SELECT BankId FROM TblPaymentType WHERE PaymentName LIKE '%" & "›Ì“«" & "%'"
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
        If Not rsDummy.EOF Then
            DcboBankName.BoundText = val(rsDummy!BankID & "")
        End If
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DcboBox.Enabled = False
        Me.DCVendor.Enabled = False
        Me.DtpChequeDueDate.Enabled = True
    Else

        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = False
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
        Fg_Journal.Clear flexClearScrollable, flexClearEverything
        Fg_Journal.rows = 3
          
      '  VSFlexGrid1.Visible = False

    ElseIf Me.CboPaymentType1.ListIndex = 1 Then

        Fg_Journal.Visible = False
        VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
        VSFlexGrid1.rows = 3
        VSFlexGrid1.Visible = True
    End If

End Sub

Private Sub CboPaymentType1_Click()
    CboPaymentType1_Change
End Sub

Private Sub Cmd_Click(Index As Integer)
    '   On Error GoTo ErrTrap

Dim Msg As String
        
            If (Index = 1 Or Index = 4) And zatcaStatus = 1 Then
                    Msg = "·« Ì„ﬂ‰  ⁄œÌ· «Ê Õ–› «Ì „” ‰œ Ì„ﬂ‰ﬂ ⁄„· „” ‰œ ⁄ﬂ”Ì ›ﬁÿ"
                        Msg = Msg & CHR(13) & ""
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
            End If
        

      If mZakamsg <> "" Then
            
        MsgBox mZakamsg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If
    
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            DcCostCenter.text = ""
            dcproject.text = ""
            
            ' XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
        
            Me.DCboUserName.BoundText = user_id
            '        XPDtbTrans.SetFocus
            Fg_Journal.Visible = False
            'VSFlexGrid1.Visible = False
DefaultInvoicetype.ListIndex = SystemOptions.DefaultInvoicetype
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.rows = 3
            Fg_Journal.Enabled = True
          
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.rows = 3
            VSFlexGrid1.Enabled = True
          
            DtpChequeDueDate.value = Date
            setfoxy
            CBoBasedON.ListIndex = 0
            CboPaymentType1.ListIndex = 0
            Me.dcBranch.BoundText = branch_id
            ComResid(1).value = True
          
        Case 1
            If ChekClodePeriod(XPDtbTrans.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
        
            
            
            If SystemOptions.banks_Accounts3 = True Then
                If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–… «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
    
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
            Fg_Journal.rows = Fg_Journal.rows + 1
            Fg_Journal.Enabled = True
            VSFlexGrid1.rows = VSFlexGrid1.rows + 1
            VSFlexGrid1.Enabled = True
        
            CuurentLogdata

        Case 2
            If ChekClodePeriod(XPDtbTrans.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
              If checkCustomerdata(val(Me.DCVendor.BoundText), val(XPTxtValView), val(DefaultInvoicetype.ListIndex), Dccurrency.text, Export) = False Then Exit Sub
            If CBoBasedON.ListIndex > 0 And Trim(TXT_order_no.text) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify NO For"
                Else
                    Msg = "Õœœ —ﬁ„ "
                End If

                Msg = Msg & "  " & CBoBasedON.text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TXT_order_no.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText
            DcboBox_Change
            DcboBankName_Change
            If val(Me.CboPayMentType.ListIndex) = 2 Then
                DCVendor_Click (0)
            End If
            SaveData "", ""
           
        Case 3
            Undo

        Case 4
            If ChekClodePeriod(XPDtbTrans.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmProjectSearch.C1Tab1 = 6
            FrmProjectSearch.Caption = "»ÕÀ «·›Ê« Ì— «·Œœ„Ì…"
            FrmProjectSearch.show vbModal

            '            Load FrmNotesSearch
            '            FrmNotesSearch.SearchType = 8063
            '            FrmNotesSearch.Show vbModal

        Case 6
            Unload Me

        Case 7
            ViewDataList

        Case 8
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report TxtSerial.text, DCVendor.text

        Case 9
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_Cheque TxtChequeNumber.text, get_Cheque_report_no(val(DcboBankName.BoundText)), TxtSerial.text

        Case 10
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc TxtSerial.text, , 8063, , , TxtSerial1.text
        Case 11
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report2 TxtSerial.text, DCVendor.text

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

    xReport.ParameterFields(5).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 1, 2)
    xReport.ParameterFields(6).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 4, 2)
    xReport.ParameterFields(7).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 9, 2)

    xReport.ParameterFields(8).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 1, 2)
    xReport.ParameterFields(9).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 4, 2)
    xReport.ParameterFields(10).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 9, 2)
    xReport.ParameterFields(11).AddCurrentValue CStr(txtto.text)
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

Function print_report(Optional NoteSerial As String, Optional VendorName As String)
    
    Dim MySQL          As String
    Dim RsData         As New ADODB.Recordset
    Dim xApp           As New CRAXDRT.Application
    Dim xReport        As CRAXDRT.Report
    Dim CViewer        As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName    As String
    Dim Msg            As String

    SaveQRCode6 "notes_all", "NoteID", val(XPTxtID.text), TxtSerial1.text, (XPDtbTrans.value), _
       (XPTxtValView.text), Picture1, 0, (txtVatTotal.text), (XPTxtValView.text), val(dcBranch.BoundText)

    'MySQL = "Select * From Expanses_Order  where noteserial='" & NoteSerial & "'"
    If CboPaymentType1.ListIndex = 0 Then
        '  MySQL = "SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial, dbo.notes_all.Note_Value, dbo.notes_all.BankID, "
        '  MySQL = MySQL & "   dbo.notes_all.ChqueNum, dbo.notes_all.DueDate, dbo.notes_all.UserID, dbo.notes_all.Remark, dbo.notes_all.ExpensesID, dbo.notes_all.BoxID,"
        '  MySQL = MySQL & "  dbo.TblUsers.UserName, dbo.TblBoxesData.BoxName, dbo.BanksData.BankName, dbo.notes_all.too, dbo.Notes.Note_Value AS [Sub-value],"
        '  MySQL = MySQL & "  dbo.Notes.note_value_by_characters AS sub_note_value_by_char, dbo.Notes.Remark AS sub_remark, dbo.ExpensesType.Name AS Sub_expenses_name,"
        '  MySQL = MySQL & "  dbo.Notes.NoteType AS sub_notetype, dbo.notes_all.note_value_by_characters, dbo.notes_all.general_des, dbo.notes_all.NoteSerial1, dbo.notes.ExpensesRemark"
        '  MySQL = MySQL & "  ,dbo.ExpensesType.Namee FROM         dbo.ExpensesType RIGHT OUTER JOIN"
        '  MySQL = MySQL & "  dbo.Notes ON dbo.ExpensesType.ID = dbo.Notes.ExpensesID LEFT OUTER JOIN"
        '  MySQL = MySQL & "  dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID FULL OUTER JOIN"
        '  MySQL = MySQL & "  dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID FULL OUTER JOIN"
        '  MySQL = MySQL & "  dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID FULL OUTER JOIN"
        '  MySQL = MySQL & "  dbo.TblUsers ON dbo.Notes.UserID = dbo.TblUsers.UserID"
        '  MySQL = MySQL & "  WHERE     (dbo.Notes.NoteType = 85) AND (NOT (dbo.ExpensesType.Name IS NULL))  and  dbo.Notes.noteserial='" & NoteSerial & "'"
        '     MySQL = "SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial, dbo.notes_all.Note_Value, dbo.notes_all.BankID, "
        'MySQL = MySQL & " dbo.notes_all.ChqueNum, dbo.notes_all.DueDate, dbo.notes_all.UserID, dbo.notes_all.Remark, dbo.notes_all.ExpensesID, dbo.notes_all.BoxID,"
        'MySQL = MySQL & " dbo.TblUsers.UserName, dbo.TblBoxesData.BoxName, dbo.BanksData.BankName, dbo.notes_all.too, dbo.Notes.Note_Value AS [Sub-value],"
        'MySQL = MySQL & " dbo.Notes.note_value_by_characters AS sub_note_value_by_char, dbo.Notes.Remark AS sub_remark, dbo.Notes.NoteType AS sub_notetype,"
        'MySQL = MySQL & " dbo.notes_all.note_value_by_characters, dbo.notes_all.general_des, dbo.notes_all.NoteSerial1, dbo.Notes.ExpensesRemark, dbo.Notes.ExpensesID AS Expr1,"
        'MySQL = MySQL & " dbo.TblRevenuesTypes.RevenuesName , dbo.TblRevenuesTypes.RevenuesNamee"
        'MySQL = MySQL & ", dbo.Notes.[Count], dbo.Notes.price , dbo.Notes.Discount   FROM         dbo.Notes INNER JOIN"
        'MySQL = MySQL & " dbo.TblRevenuesTypes ON dbo.Notes.ExpensesID = dbo.TblRevenuesTypes.RevenuesID LEFT OUTER JOIN"
        'MySQL = MySQL & " dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID FULL OUTER JOIN"
        'MySQL = MySQL & " dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID FULL OUTER JOIN"
        'MySQL = MySQL & " dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID FULL OUTER JOIN"
        'MySQL = MySQL & " dbo.TblUsers ON dbo.Notes.UserID = dbo.TblUsers.UserID"
        MySQL = " SELECT     dbo.notes_all.QrCodeImage , dbo.notes_all.PayAmount, dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial, dbo.notes_all.Note_Value,notes_all.NewNO,notes_all.ComResid, dbo.notes_all.BankID,DOUBLE_ENTREY_VOUCHERS.VAT,DOUBLE_ENTREY_VOUCHERS.Vatyo ,"
        MySQL = MySQL & "                      dbo.notes_all.ChqueNum, dbo.notes_all.DueDate, dbo.notes_all.UserID, dbo.notes_all.Remark, dbo.notes_all.ExpensesID, dbo.notes_all.BoxID,notes_all.Vendor,"
        MySQL = MySQL & "                      dbo.TblUsers.UserName, dbo.TblBoxesData.BoxName, dbo.BanksData.BankName, dbo.notes_all.too, dbo.Notes.Note_Value AS [Sub-value],"
        MySQL = MySQL & "                      dbo.Notes.note_value_by_characters AS sub_note_value_by_char, dbo.Notes.Remark AS sub_remark, dbo.Notes.NoteType AS sub_notetype,Notes.des2,Notes.ContItem,Notes.PurchaseOrderNo,Notes.LocationName,Notes.RentType,Notes.UnitNo,"
        MySQL = MySQL & "                      dbo.notes_all.note_value_by_characters, dbo.notes_all.general_des, dbo.notes_all.NoteSerial1, dbo.Notes.ExpensesRemark, dbo.Notes.ExpensesID AS Expr1,"
        MySQL = MySQL & "                      dbo.TblRevenuesTypes.RevenuesName, dbo.TblRevenuesTypes.RevenuesNamee, dbo.Notes.[Count], dbo.Notes.price, dbo.Notes.discount, dbo.notes_all.CarId,"
        MySQL = MySQL & "                      dbo.TblCarsData.BoardNO, dbo.DOUBLE_ENTREY_VOUCHERS.Carid AS CaridDet, TblCarsData_1.BoardNO AS BoardNODet"
 
        MySQL = MySQL & "        ,dbo.notes_all.NoteSerial,notes_all.TotalFines,notes_all.RequestNo,notes_all.ContractNo,notes_all.Noteseril2"
        MySQL = MySQL & "   ,dbo.TblCustemers.Cus_mobile , dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.E_mail,TblCustemers.FullCode"
        MySQL = MySQL & "  ,dbo.TblCustemers.Address ,TblCustemers.CustGID, TblCustemers.VATNO,notes_all.TotalFines ,Notes.des2,Notes.NoteID as Serial,TblCustemers.CusName,TblCustemers.CusNamee,TblCustemers.AddressE ,"
        MySQL = MySQL & "  projects.Project_nameE,projects.Project_name ,"
        MySQL = MySQL & "                      TblBranchesData.Company_Arabic_Name,TblBranchesData.Company_Name_Eng,"
        MySQL = MySQL & "                         TblBranchesData.CityName as CompanyCityName,TblBranchesData.CitySubdivisionName,TblBranchesData.PostalZone CompanyPostalZone,TblBranchesData.StreetName CompanyStreetName,TblBranchesData.VATRegNo CompanyVATRegNo,TblBranchesData.Company_Comment CompanyRegNo"
        MySQL = MySQL & "  ,dbo.TblCustemers.Address ,TblCustemers.CustGID, TblCustemers.VATNO,notes_all.TotalFines ,Notes.des2,Notes.NoteID as Serial,TblCustemers.CusName,TblCustemers.CusNamee,TblCustemers.AddressE "
        MySQL = MySQL & ",notes_all.DueDate,TblCountriesData.CountryName,TblCountriesData.ECountryName,TblCountriesGovernments.GovernmentName"
        MySQL = MySQL & "   FROM         dbo.TblBoxesData RIGHT OUTER JOIN"
        MySQL = MySQL & "                      dbo.TblCarsData TblCarsData_1 RIGHT OUTER JOIN"
        MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS ON TblCarsData_1.id = dbo.DOUBLE_ENTREY_VOUCHERS.Carid RIGHT OUTER JOIN"
        MySQL = MySQL & "                       dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
        MySQL = MySQL & "                      dbo.TblRevenuesTypes ON dbo.Notes.ExpensesID = dbo.TblRevenuesTypes.RevenuesID ON dbo.TblBoxesData.BoxID = dbo.Notes.BoxID LEFT OUTER JOIN"
        MySQL = MySQL & "                      dbo.TblUsers ON dbo.Notes.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
        MySQL = MySQL & "                      dbo.TblCarsData RIGHT OUTER JOIN"
        MySQL = MySQL & "                      dbo.notes_all ON dbo.TblCarsData.id = dbo.notes_all.CarId ON dbo.Notes.notes_all = dbo.notes_all.NoteID LEFT OUTER JOIN"
        MySQL = MySQL & "                      dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID"
        MySQL = MySQL & "                       Left OUTER JOIN TblCustemers On  notes_all.CusID = TblCustemers.CusID"
        MySQL = MySQL & "                       Left OUTER JOIN projects On  notes_all.project_Expensen_account = projects.REVENUE_account and Isnull(notes_all.project_Expensen_account,'') <> ''"
        MySQL = MySQL & "                       Left OUTER JOIN TblBranchesData On  notes_all.branch_no = TblBranchesData.branch_id"
        
        MySQL = MySQL & "                       Left Outer join TblCountriesData On TblCustemers.CountryID =TblCountriesData.CountryID "
        MySQL = MySQL & "                       Left Outer join TblCountriesGovernments On TblCustemers.GovernmentID =TblCountriesGovernments.GovernmentID "
 
        MySQL = MySQL & "  WHERE     (dbo.Notes.NoteSerial = '" & NoteSerial & "') AND (dbo.notes_all.NoteType = 85)  AND (NOT (dbo.Notes.ExpensesRemark IS NULL))"
        MySQL = MySQL & "    AND (dbo.Notes.NoteSerial1 = " & TxtSerial1 & ")"
 
        'MySQL = MySQL & "  WHERE     (dbo.Notes.NoteType = 85) AND (NOT (dbo.TblRevenuesTypes.RevenuesNamee IS NULL))  and  dbo.Notes.noteserial='" & NoteSerial & "'"
         
        MySQL = MySQL & "             and (DOUBLE_ENTREY_VOUCHERS.Account_code Not In ( Select AccCir  from TblSettsReqLimK) and  DOUBLE_ENTREY_VOUCHERS.Account_code Not In ( Select accDep  from TblSettsReqLimK))"
        MySQL = MySQL & "   order BY Notes.NoteID"
    Else
        MySQL = "SELECT    dbo.notes.QrCodeImage , dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.BankID, dbo.Notes.ChqueNum, dbo.Notes.DueDate, "
        MySQL = MySQL & "   dbo.Notes.CusID, dbo.Notes.BoxID, dbo.Notes.Note_Value, dbo.Notes.note_value_by_characters, dbo.Notes.Remark AS sub_remark, dbo.ACCOUNTS.Account_Name,"
        MySQL = MySQL & "  dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS [sub-value],DOUBLE_ENTREY_VOUCHERS.VAT,DOUBLE_ENTREY_VOUCHERS.Vatyo,"
        MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS expenses_remark, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit,Notes.des2,Notes.NoteID as Serial"
        MySQL = MySQL & "  FROM         dbo.Notes INNER JOIN"
        MySQL = MySQL & "  dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
        MySQL = MySQL & "  dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
        MySQL = MySQL & "   WHERE     (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND (dbo.Notes.NoteSerial = " & NoteSerial & ")"
        MySQL = MySQL & "   order BY Notes.NoteID"
    End If

    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If

    If CboPaymentType1.ListIndex = 0 Then
        
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\New Reports\" & "ServiceInvoice.rpt"
        Else
            StrFileName = App.path & "\Reports\New Reports\" & "ServiceInvoice.rpt"
        End If

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Reports\New Reports\" & "FinancialInvoiceAccounts.rpt"
        Else
            StrFileName = App.path & "\Reports\New Reports\" & "FinancialInvoiceAccountse.rpt"
        End If

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
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(6).AddCurrentValue VendorName


  Dim i As Integer
    For i = 1 To xReport.FormulaFields.count
         Select Case xReport.FormulaFields.Item(i).Name
         Case "{@VATRegNo}"
             If SystemOptions.VATNoAccordActivity = False Then
                ' xReport.FormulaFields.Item(i).text = "'" & StrRS!VATRegNo & "'"
             Else
                 xReport.FormulaFields.Item(i).text = "'" & GetRegVATNo(branch_id) & "'"
             End If
             Case "{@HijriDate}"
                xReport.FormulaFields.Item(i).text = "'" & GetHijriDate(Date) & "'"
             Case "{@MonthName}"
                xReport.FormulaFields.Item(i).text = "'" & GetMonthName(Month(XPDtbTrans.value)) & "'"
             Case "{@TotalEnglish}"
                xReport.FormulaFields.Item(i).text = "'" & WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 1) & "'"
             Case "{@TotalEnglish2}"
                xReport.FormulaFields.Item(i).text = "'" & WriteNo(Format(Me.XPTxtVal2.text, "0.00"), 0, True, ".", , 1) & "'"
             Case "{@TotalArabic}"
                xReport.FormulaFields.Item(i).text = "'" & WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 0) & "'"
             Case "{@TotalArabic2}"
                xReport.FormulaFields.Item(i).text = "'" & WriteNo(Format(Me.XPTxtVal2.text, "0.00"), 0, True, ".", , 0) & "'"
                
         End Select
     Next i



'
'   If SystemOptions.UserInterface = ArabicInterface Then
'        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 0)
'        Me.lblValue2.Caption = WriteNo(Format(Me.XPTxtVal2.text, "0.00"), 0, True, ".", , 0)
'
'    Else
'
'        'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0, , , , 1)
'        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 1)
'        Me.lblValue2.Caption = WriteNo(Format(Me.XPTxtVal2.text, "0.00"), 0, True, ".", , 1)
'
'    End If
'
'
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function


Function print_report2(Optional NoteSerial As String, Optional VendorName As String, Optional ByVal mIndex As Long = 0)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim StrRS As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim s As String
    
   

    'MySQL = "Select * From Expanses_Order  where noteserial='" & NoteSerial & "'"
    
     rs("FromDate").value = txtFromDate.value
       rs("FromDateH").value = Me.txtFromDateH.value
       
   rs("ToDate").value = txtToDate.value
       rs("ToDateh").value = Me.txtToDateH.value

    s = " SELECT dbo.notes_all.NoteID,"
     
    
     
    s = s & "         CAST(DAY(notes_all.NoteDate) AS NVARCHAR(2))   + '/' + CAST(MONTH(notes_all.NoteDate) AS NVARCHAR(2)) + '/'  + CAST(YEAR(notes_all.NoteDate) AS NVARCHAR(4))  AS  NoteDate2,"
    's = s & "      CAST(notes_all.NoteDate AS NVARCHAR(10)) as    NoteDate,"
    s = s & "        dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description,Notes.DailyMonthlyValue,notes_all.NoteDate,"
    s = s & "        dbo.notes_all.NoteType,notes_all.FromDate,notes_all.FromDateH,notes_all.ToDate,notes_all.ToDateH,"
    s = s & "        dbo.notes_all.NoteSerial,notes_all.TotalFines,notes_all.RequestNo,notes_all.ContractNo,notes_all.Noteseril2,"
    s = s & "        dbo.notes_all.Note_Value,"
    s = s & "        dbo.notes_all.BankID,"
    s = s & "        dbo.notes_all.ChqueNum,"
    s = s & "        dbo.notes_all.DueDate,"
    s = s & "        dbo.notes_all.UserID,"
    s = s & "        dbo.notes_all.Remark,"
    s = s & "        dbo.notes_all.ExpensesID,"
    s = s & "        dbo.notes_all.BoxID,"
    s = s & "        dbo.TblUsers.UserName,"
    s = s & "        dbo.TblBoxesData.BoxName,"
    s = s & "        dbo.BanksData.BankName,"
    s = s & "        dbo.notes_all.too,dbo.Notes.des2,"
    s = s & "        dbo.Notes.Note_Value              AS [Sub-value],"
    s = s & "        dbo.Notes.note_value_by_characters AS sub_note_value_by_char,"
    s = s & "        dbo.Notes.note_v_by_char_WithoutVat ,"
    
    s = s & "        dbo.Notes.Remark                  AS sub_remark,"
    s = s & "        dbo.Notes.NoteType                AS sub_notetype,"
    s = s & "        dbo.notes_all.note_value_by_characters,"
    s = s & "        dbo.notes_all.general_des,"
    s = s & "        dbo.notes_all.NoteSerial1,notes_all.TotalFines,"
    s = s & "        dbo.Notes.ExpensesRemark,"
    s = s & "        dbo.Notes.ExpensesID              AS Expr1,"
    s = s & "        dbo.TblRevenuesTypes.RevenuesName,"
    s = s & "        dbo.TblRevenuesTypes.RevenuesNamee,"
    s = s & "        dbo.Notes.[Count],"
    s = s & "        dbo.Notes.price,"
    s = s & "        dbo.Notes.discount,"
    s = s & "        dbo.notes_all.CarId,"
    s = s & "        dbo.Notes.des2 as BoardNO,"
    s = s & "        dbo.DOUBLE_ENTREY_VOUCHERS.Carid  AS CaridDet,"
    s = s & "        TblCarsData_1.BoardNO             AS BoardNODet,"
    s = s & "        Notes.MonthCount,"
    s = s & "        Notes.PurchOrderNo,"
    s = s & "        Notes.CityFromId,"
    s = s & "        Notes.CityToId,"
    s = s & "        FromCity.GovernmentName              FromCityName,"
    s = s & "        ToCity.GovernmentName ToCityName"
    s = s & " ,dbo.TblCustemers.Cus_mobile , dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.E_mail"
    s = s & " ,dbo.TblCustemers.Address , TblCustemers.VATNO"
    s = s & "               ,TblCustemers.CustGID,TblCustemers.FullCode ,"
    s = s & "               TblCustemers.CusName,TblCustemers.CusNamee,DOUBLE_ENTREY_VOUCHERS.Vat"
    s = s & "               ,TblCustemers.GovernmentID , TblCustemers.CityID"
    s = s & "               ,CustCity.CityName as CustCityName , CustG.GovernmentName CustGovernmentName"
    s = s & "               ,Notes.DailyMonthly,"
    s = s & "                      notes_all.QrCodeImage,TblCustemers.VATNO,TblCustemers.Address as AddressSupp"
    s = s & " From dbo.TblBoxesData"
    s = s & "        RIGHT OUTER JOIN dbo.TblCarsData TblCarsData_1"
    s = s & "        RIGHT OUTER JOIN dbo.DOUBLE_ENTREY_VOUCHERS"
    s = s & "             ON  TblCarsData_1.id = dbo.DOUBLE_ENTREY_VOUCHERS.Carid"
    s = s & "        RIGHT OUTER JOIN dbo.Notes"
    s = s & "             ON  dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
    
    s = s & "        LEFT OUTER JOIN dbo.TblRevenuesTypes"
    s = s & "             ON  dbo.Notes.ExpensesID = dbo.TblRevenuesTypes.RevenuesID"
    s = s & "             ON  dbo.TblBoxesData.BoxID = dbo.Notes.BoxID"
    s = s & "        LEFT OUTER JOIN dbo.TblUsers"
    s = s & "             ON  dbo.Notes.UserID = dbo.TblUsers.UserID"
    s = s & "        LEFT OUTER JOIN dbo.TblCarsData"
    s = s & "        RIGHT OUTER JOIN dbo.notes_all"
    s = s & "             ON  dbo.TblCarsData.id = dbo.notes_all.CarId"
    s = s & "             ON  dbo.Notes.notes_all = dbo.notes_all.NoteID"
    
    s = s & "             and (ExpensesRemark NOT LIKE '%ﬁÌ„… „÷«›…%' AND ISNULL(Notes.ExpensesID,0) <> 0)"
    
    s = s & "        LEFT OUTER JOIN dbo.BanksData"
    s = s & "             ON  dbo.Notes.BankID = dbo.BanksData.BankID"
    s = s & "        LEFT OUTER JOIN TblCountriesGovernments FromCity"
    s = s & "             ON  Notes.CityFromId = FromCity.GovernmentID"
    s = s & "        LEFT OUTER JOIN TblCountriesGovernments ToCity"
    s = s & "             ON  Notes.CityToId = ToCity.GovernmentID"
    s = s & "                  LEFT OUTER JOIN dbo.TblCustemers"
    s = s & "                         ON  notes_all.CusID = TblCustemers.CusID"
    s = s & "                  LEFT OUTER JOIN dbo.TblCountriesGovernments CustG"
    s = s & "             ON  TblCustemers.GovernmentID = CustG.GovernmentID"
    s = s & "                  LEFT OUTER JOIN dbo.TblCountriesGovernmentsCities CustCity"
    s = s & "             ON  TblCustemers.CityId = CustCity.CityId"
    
 '   db_createOrUpdateviewSQL "Vw_Notes", s


     
     s = s & "  WHERE     (dbo.Notes.NoteSerial = '" & NoteSerial & "') AND (dbo.notes_all.NoteType = 85)  AND (NOT (dbo.Notes.ExpensesRemark IS NULL))"
     s = s & "    AND (dbo.Notes.NoteSerial1 = " & TxtSerial1 & ")   "
     s = s & "    "
'     If mIndex = 0 Then
'        s = s & " AND (dbo.Notes.DailyMonthly) = 0"
'     ElseIf mIndex = 2 Then
'        s = s & " AND (dbo.Notes.DailyMonthly) = 1"
'     End If
'
  'MySQL = MySQL & "  WHERE     (dbo.Notes.NoteType = 85) AND (NOT (dbo.TblRevenuesTypes.RevenuesNamee IS NULL))  and  dbo.Notes.noteserial='" & NoteSerial & "'"
         
        
        
        
        
        
   

    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If

 
        If mIndex = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\New Reports\" & "ServiceInvoice2.rpt"
            Else
                StrFileName = App.path & "\Reports\New Reports\" & "ServiceInvoice2.rpt"
            End If
        ElseIf mIndex = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\New Reports\" & "ServiceInvoice3.rpt"
            Else
                StrFileName = App.path & "\Reports\New Reports\" & "ServiceInvoice3.rpt"
            End If
        ElseIf mIndex = 2 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\New Reports\" & "ServiceInvoice4.rpt"
            Else
                StrFileName = App.path & "\Reports\New Reports\" & "ServiceInvoice4.rpt"
            End If
        ElseIf mIndex = 3 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\Reports\New Reports\" & "ServiceInvoice5.rpt"
            Else
                StrFileName = App.path & "\Reports\New Reports\" & "ServiceInvoice5.rpt"
            End If
        
        End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

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

    Set StrRS = New ADODB.Recordset
    StrRS.Open "TblOptions", Cn, adOpenStatic, adLockOptimistic, adCmdTable
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
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If
    
    
    Dim i As Integer
    For i = 1 To xReport.FormulaFields.count
         Select Case xReport.FormulaFields.Item(i).Name
         Case "{@VATRegNo}"
             If SystemOptions.VATNoAccordActivity = False Then
                 xReport.FormulaFields.Item(i).text = "'" & StrRS!VATRegNo & "'"
             Else
                 xReport.FormulaFields.Item(i).text = "'" & GetRegVATNo(branch_id) & "'"
             End If
             Case "{@HijriDate}"
                xReport.FormulaFields.Item(i).text = "'" & GetHijriDate(Date) & "'"
             Case "{@MonthName}"
                xReport.FormulaFields.Item(i).text = "'" & GetMonthName(Month(XPDtbTrans.value)) & "'"
                
         End Select
     Next i



'    xReport.ParameterFields(3).AddCurrentValue user_name
'    xReport.ParameterFields(6).AddCurrentValue VendorName

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , s

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function



Public Function GetMonthName(ByVal mMonthNo As Long) As String
    
    Dim mLang As Boolean
    mLang = SystemOptions.UserInterface = ArabicInterface
    If SystemOptions.UserInterface = ArabicInterface Then
        Select Case mMonthNo
        Case 1
            GetMonthName = IIf(mLang, "Ì‰«Ì—", "january")
        Case 2
            GetMonthName = IIf(mLang, "›»—«Ì—", "February")
        Case 3
            GetMonthName = IIf(mLang, "„«—”", "March")
        Case 4
            GetMonthName = IIf(mLang, "«»—Ì·", "April")
        Case 5
            GetMonthName = IIf(mLang, "„«ÌÊ", "May")
        Case 6
            GetMonthName = IIf(mLang, "ÌÊ‰ÌÊ", "june")
        Case 7
            GetMonthName = IIf(mLang, "ÌÊ·ÌÊ", "July")
        Case 8
            GetMonthName = IIf(mLang, "√€”ÿ”", "August")
        Case 10
            GetMonthName = IIf(mLang, "”» „»—", "September")
        Case 11
            GetMonthName = IIf(mLang, "‰Ê›„»—", "Nov")
        Case 12
            GetMonthName = IIf(mLang, "œÌ”„»—", "Dec")
       End Select
    End If

End Function
Private Sub CmdAttach_Click()
    On Error Resume Next
ShowAttachments TxtSerial1, "0612201506"

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub
Sub LoadExcel()
    On Error GoTo eh
 
    If grdFiles.rows > 1 Then
        Dim i        As Integer
       
        Dim tblname  As String
        Dim FileName As String
        Dim s        As String
        ' Add ref to
        ' Microsoft Scripting Runtime
        Dim fn       As New FileSystemObject
        allExcelRows = 0
        Dim oConn  As ADODB.Connection
        Dim mrs    As ADODB.Recordset
        Dim RsData As ADODB.Recordset
        Set oConn = New ADODB.Connection
     
        Set RsData = New ADODB.Recordset
        For i = 1 To grdFiles.rows - 1
           
            FileName = grdFiles.TextMatrix(i, grdFiles.ColIndex("FileName"))
            oConn.CursorLocation = adUseClient
            oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
               & "Data Source='" & FileName & "';" _
               & "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
            '& "Extended Properties=""Excel 8.0;"""
            Set mrs = oConn.OpenSchema(adSchemaTables)
            If Not mrs.EOF Then
                tblname = mrs.Fields("table_name").value
                RsData.CursorLocation = adUseClient
                RsData.Open "Select *   from [" & tblname & "]", oConn, adOpenKeyset, adLockReadOnly
                RsData.MoveLast
                grdFiles.TextMatrix(i, grdFiles.ColIndex("Rows")) = RsData.RecordCount
                allExcelRows = allExcelRows + val(RsData.RecordCount)
                RsData.Close
            End If
            
            Dim FileNameOnly As String
            Dim mfile        As File
            
            Set mfile = fn.GetFile(FileName)
            grdFiles.TextMatrix(i, grdFiles.ColIndex("Name")) = mfile.Name
            Set mfile = Nothing
            mrs.Close
            oConn.Close
         
        Next
    End If
    Exit Sub
eh:
    MsgBox Err.Description
End Sub

Private Sub cmdLoadFile_Click()
    On Error GoTo eh

    If grdFiles.rows = 1 Then
        MsgBox "«Œ — „·› «Ê·«"
        Exit Sub
    End If

    'ExportToExcel Me, Grd, "TT", , "grdExcel"
    tmpGrd.rows = 1
    ' lbl(32).Visible = True
    Dim i             As Long
    Dim s             As String
  
    '    *****OLD CODE************
    '      grdExcel(0).Rows = 1
    '     FromExcel grdExcel(0), tmpGrd, Me, , , txtFile.Text, "TblEmployee"
    '       SaveItemsExcelMeth3
    '    *******************
    '*****New CODE************
    ' grdExcel(0).Rows = 1
    ' FromExcel GrdExcel(0), tmpGrd, Me, , , txtFile.Text, "TblEmployee"     Dim i As Integer
    Dim RsData        As New ADODB.Recordset
    Dim FileName      As String
    Dim moConn        As New ADODB.Connection
    Dim mrs           As ADODB.Recordset
    Dim tblname       As String
    Dim shortFileName As String
    lblTime.Visible = True
    Me.Enabled = False
    moConn.CursorLocation = adUseClient
    Dim rsCheck As New ADODB.Recordset

    For i = 1 To grdFiles.rows - 1
        FileName = grdFiles.TextMatrix(i, grdFiles.ColIndex("FileName"))
        shortFileName = grdFiles.TextMatrix(i, grdFiles.ColIndex("Name"))
        
        moConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source='" & FileName & "';" & "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
        Set mrs = moConn.OpenSchema(adSchemaTables)

        If Not mrs.EOF Then
            tblname = mrs.Fields("table_name").value
            RsData.CursorLocation = adUseClient
            RsData.Open "Select *   from [" & tblname & "]", moConn, adOpenKeyset, adLockReadOnly
            Dim RowID          As Integer
           
            Dim isREcoredSaved As Boolean
            Dim strQuery       As String
            Dim OLDSec         As Long
            Dim Secondes       As Long

            Dim AllSec         As Long
            lblTime.Visible = True
            RowID = 0
            AllFinshedRows = 0
            Dim currentRows As Long
            Dim AllFileRows As Long
            AllFileRows = val(grdFiles.TextMatrix(i, grdFiles.ColIndex("Rows")))

            Do While Not RsData.EOF
                RowID = RowID + 1
                AllFinshedRows = AllFinshedRows + 1
                currentRows = currentRows + 1
                ' lbl(32).Caption = "F[" & currentRows & "]>[" & AllFileRows & "] A[" & AllFinshedRows & "]>[" & AllFileRows & "]"
                DoEvents
                strQuery = "SELECT Count(*) cnt "
                strQuery = strQuery & "From notes_all "
                strQuery = strQuery & "WHERE ExcelFile = '" & shortFileName & "' "
                strQuery = strQuery & " AND NoteType = 85 "
                strQuery = strQuery & "  AND ExcelRow =  " & RowID & " ;"
                ' rsCheck.CursorLocation = adUseClient
                rsCheck.Open strQuery, Cn, adOpenForwardOnly, adLockReadOnly
                isREcoredSaved = rsCheck!cnt > 0
                rsCheck.Close

                '*********************
                If Not isREcoredSaved Then
                    startTime = Now

                    If Trim(RsData.Fields("«· «—ÌŒ").value & "") <> "" Then
                        SaveItemsExcelMeth_New RsData, RowID, shortFileName
                    End If

                    OLDSec = AllSec
                    Secondes = DateDiff("s", startTime, Now)

                    AllSec = ((allExcelRows - AllFinshedRows) * Secondes)

                    If AllSec = 0 Then
                        AllSec = OLDSec
                    End If

                    lblTime.Caption = StringDotFormat("{0} of {1} Recored(s)  Estimated Time : {2} ", AllFinshedRows, allExcelRows, GetTimeHour(AllSec))
                End If

                '*********************
                RsData.MoveNext
            Loop

            RsData.Close
        End If

        mrs.Close
        moConn.Close

    Next

    Me.Enabled = True
    lblTime.Visible = False
    grdFiles.rows = 1
    MsgBox " „ Õ›Ÿ «·Õ—ﬂ« "
    Exit Sub
eh:
    Me.Enabled = True
    lblTime.Visible = False
    MsgBox Err.Description
End Sub

Private Sub SaveItemsExcelMeth_New(excelRS As ADODB.Recordset, RowID, mFileName)
    Dim i                   As Long
    Dim s                   As String
    Dim rsDummy             As New ADODB.Recordset
    Dim rsDummy2            As New ADODB.Recordset
    Dim mCode               As String
    Dim mGroupID            As Long
    Dim mUnitId             As Long
    Dim mUnitPurPrice       As Double
    Dim mUnitSalesPrice     As Double
    Dim mRatePur            As Double
    Dim mRateSale           As Double
    Dim mNewCode            As String
    Dim mMaxId              As Long
    Dim mUnitName           As String
    Dim mName               As String
    Dim mbarCode            As String
    Dim mUnitWholeSalePrice As Double

    Dim rsDummySupp         As New ADODB.Recordset
    Dim mIndex              As Long
    Dim mbranch_id          As Long
    Dim mBoxID              As Long
    Dim mCusId              As Long
    Dim mNoteDate           As Date

    mIndex = 0

    Dim rsDummyUnit        As New ADODB.Recordset
    Dim mMinSelingPrice    As Double
    Dim mMaxSelingPrice    As Double
    Dim mSelingPriceDestr  As Double
    Dim mDefaultSupplier   As String
    Dim mDefaultSupplierID As Integer

    Dim mBranch_Name       As String
    Dim mNoteCashingType   As String
    Dim mBoxName           As String

    Dim mCusName           As String
    Dim mAccountName       As String
    Dim mAccountCode       As String
    Dim mgeneral_des       As String
    Dim mPrice             As Double
    Dim mFullCode          As String
    '--Select replace (@mm,'S

    '  For i = 1 To grdExcel(mIndex).Rows - 1
    ' Do While Not excelRS.EOF
    ' If i = grdExcel(mIndex).Rows Then GoTo NextRow
    mCode = excelRS.Fields("—ﬁ„ «·⁄„Ì·").value & ""  'Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("Fullcode")))
    mBranch_Name = excelRS.Fields("«”„ «·›—⁄").value & ""  'Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("Branch_Name")))
    mbranch_id = 0 'val(Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("Branch_ID"))))
    mNoteDate = excelRS.Fields("«· «—ÌŒ").value & ""  'Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("NoteDate")))
    mPrice = Round(val(excelRS.Fields("«·„»·€").value & ""), Decimal_Places)  ' val(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("price")))
    mNoteCashingType = excelRS.Fields("‰Ê⁄ «·›« Ê—…").value & ""  'Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("NoteCashingType")))
        
    mBoxName = excelRS.Fields("«”„ «·»«∆⁄").value & ""  'Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("BoxName")))
    mBoxID = val(excelRS.Fields("—ﬁ„ «·»«∆⁄").value & "")  'val(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("BoxID")))
    '

    mCusName = excelRS.Fields("«”„ «·⁄„Ì·").value & ""  'Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("CusName")))
    mCusId = val(excelRS.Fields("—ﬁ„ «·⁄„Ì·").value & "")   ' val(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("CusID")))
    mAccountName = excelRS.Fields("«”„ «·Õ”«»").value & ""  'Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("AccountName")))
    mgeneral_des = excelRS.Fields("«·‘—Õ").value & ""  'Trim(grdExcel(mIndex).TextMatrix(i, grdExcel(mIndex).ColIndex("general_des")))
        
    If mNoteCashingType = "‰ﬁœÌ" Or mNoteCashingType = "‰ﬁœÏ" Or mNoteCashingType = "‰ﬁœ«" Then
        CboPayMentType.ListIndex = 0
    
    ElseIf mNoteCashingType = "„œÏ" Or mNoteCashingType = "„œÌ" Then
        CboPayMentType.ListIndex = 4
    ElseIf mNoteCashingType = "›Ì“" Or mNoteCashingType = "›Ì“«" Then
        CboPayMentType.ListIndex = 5
    Else
        CboPayMentType.ListIndex = 2
    End If
        
        

  
    txt_general_des = mgeneral_des
    Set rsDummyUnit = New ADODB.Recordset
    s = "Select * from TblBranchesData Where Branch_Name Like '%" & Trim(mBranch_Name) & "%' Or  Branch_Code = '" & Trim(mBranch_Name) & "'"
                 
    rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummyUnit.EOF Then
        mbranch_id = val(rsDummyUnit!branch_id & "")
    Else
        mbranch_id = Current_branch
    End If
    rsDummyUnit.Close
    
    If CboPayMentType.ListIndex = 0 Then
        s = " SELECT     BoxId,BoxName from tblBoxesData Where BoxName Like '%" & Trim(mBoxName) & "%'"
         
        rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummyUnit.EOF Then
            mBoxID = val(rsDummyUnit!BoxID & "")
        End If
         
        rsDummyUnit.Close
    End If
    If CboPayMentType.ListIndex > 3 Then
        CboPayMentType_Change
    End If
    If CboPayMentType.ListIndex <> 0 And CboPayMentType.ListIndex < 4 Then
        s = ""
        If mCode <> "" Then
            s = " SELECT     CusID,CusName from TblCustemers Where  FullCode = '" & Trim(mCode) & "'"
        Else
            s = " SELECT     CusID,CusName from TblCustemers Where CusName = N'%" & Trim(mCusName) & "%' "
        
        End If

        Set rsDummyUnit = New ADODB.Recordset
        rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly
        If rsDummyUnit.EOF Then
            rsDummyUnit.Close
        
            s = " SELECT     CusID,CusName from TblCustemers Where CusName Like N'%" & Trim(mCusName) & "%' "
            rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly
        End If
        If Not rsDummyUnit.EOF Then
            mCusId = val(rsDummyUnit!CusID & "")
        End If
    
        rsDummyUnit.Close
     
    End If
    s = " SELECT     Account_Code,Account_Name from Accounts Where Account_Name = '" & Trim(mAccountName) & "'"
    
    rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummyUnit.EOF Then
        mAccountCode = Trim(rsDummyUnit!Account_code & "")
    End If
    
    rsDummyUnit.Close
    mIsSaveFromExcel = True
    
    '    Set rsDummySupp = New ADODB.Recordset
    '    s = "SELECT CusID FROM TblCustemers Where (CusName Like '%" & Trim(mDefaultSupplier) & "%'     Or CusNamee Like '%" & Trim(mDefaultSupplier) & "%')"
    '    rsDummySupp.Open s, Cn, adOpenStatic, adLockReadOnly
    '    If Not rsDummySupp.EOF Then
    '        mDefaultSupplierID = val(rsDummySupp!CusID & "")
    '    End If
    '
    TxtModFlg.text = "N"
    isCmdCreateTransa = True
    
    XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
    '  Me.TxtNoteSerial.Text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
    Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=85"))
    Me.DCboUserName.BoundText = user_id
    '        XPDtbTrans.SetFocus
              
    XPDtbTrans.value = mNoteDate
    dcBranch.BoundText = mbranch_id
    If mCusId <> 0 Then
        DCVendor.BoundText = mCusId
        'DCVendor_Click 0
        If DCVendor.BoundText <> "" Then
            If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DCVendor.BoundText))
            End If
        End If
    End If
    If mBoxID <> 0 Then
        DcboBox.BoundText = mBoxID
        DcboBox_Change
    End If
     If CboPayMentType.ListIndex > 3 Then
        CboPayMentType_Change
    End If
    dcBranch.BoundText = mbranch_id
    '               Dcbranch.BoundText = mbranch_id
    '               Dcbranch.BoundText = mbranch_id
        
    Fg_Journal.rows = 1
    Fg_Journal.rows = 2
    Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("AccountName")) = mAccountName
    Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("AccountCode")) = mAccountCode
    Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("Count")) = 1
    Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("price")) = mPrice
    'Fg_Journal_AfterEdit 1, Fg_Journal.ColIndex("AccountName")
    Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("ExpensesID")) = get_Revenue_id(mAccountCode)
    Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("LineNo1")) = setfoxy_Line
    Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("Order_No")) = TXT_order_no.text
    Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("Count")) = 1
    Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("des")) = mgeneral_des

    'ddVATExp 2
    If ComResid(1).value Then
            AddVATExp 1
    End If
    
    ReLineGrid
    Fg_Journal.Row = 1
    
    If ComResid(1).value Then
                    AddVATExp 1
                    AddVATExp 1
    End If
   
    
    setfoxy
    TxtSerial1.text = ""
    TxtSerial = ""
    TxtNoteSerial = ""
    TxtNoteSerial1 = ""
    SaveData RowID, mFileName
    If mBranch_Name = "" Then GoTo NextRow
NextRow:
    'excelRS.MoveNext
    '
    '    Loop

    mIsSaveFromExcel = False
End Sub



Private Sub SaveItemsExcelMeth3()
    Dim i                   As Long
    Dim s                   As String
    Dim rsDummy             As New ADODB.Recordset
    Dim rsDummy2            As New ADODB.Recordset
    Dim mCode               As String
    Dim mGroupID            As Long
    Dim mUnitId             As Long
    Dim mUnitPurPrice       As Double
    Dim mUnitSalesPrice     As Double
    Dim mRatePur            As Double
    Dim mRateSale           As Double
    Dim mNewCode            As String
    Dim mMaxId              As Long
    Dim mUnitName           As String
    Dim mName               As String
    Dim mbarCode            As String
    Dim mUnitWholeSalePrice As Double

    Dim rsDummySupp         As New ADODB.Recordset
    Dim mIndex              As Long
    Dim mbranch_id          As Long
    Dim mBoxID              As Long
    Dim mCusId              As Long
    Dim mNoteDate           As Date

    mIndex = 0

    Dim rsDummyUnit        As New ADODB.Recordset
    Dim mMinSelingPrice    As Double
    Dim mMaxSelingPrice    As Double
    Dim mSelingPriceDestr  As Double
    Dim mDefaultSupplier   As String
    Dim mDefaultSupplierID As Integer

    Dim mBranch_Name       As String
    Dim mNoteCashingType   As String
    Dim mBoxName           As String

    Dim mCusName           As String
    Dim mAccountName       As String
    Dim mAccountCode       As String
    Dim mgeneral_des       As String
    Dim mPrice             As Double
    Dim mFullCode          As String
    '--Select replace (@mm,'S

    For i = 1 To GrdExcel(mIndex).rows - 1
        If i = GrdExcel(mIndex).rows Then GoTo NextRow
        mCode = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("Fullcode")))
        mBranch_Name = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("Branch_Name")))
        mbranch_id = val(Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("Branch_ID"))))
        mNoteDate = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("NoteDate")))
        mPrice = val(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("price")))
        mNoteCashingType = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("NoteCashingType")))
        mBoxName = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("BoxName")))
        mBoxID = val(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("BoxID")))
        mCusName = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("CusName")))
        mCusId = val(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("CusID")))
        mAccountName = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("AccountName")))
        mgeneral_des = Trim(GrdExcel(mIndex).TextMatrix(i, GrdExcel(mIndex).ColIndex("general_des")))
        If mNoteCashingType = "‰ﬁœÌ" Or mNoteCashingType = "‰ﬁœÏ" Or mNoteCashingType = "‰ﬁœ«" Then
            CboPayMentType.ListIndex = 0
        Else
            CboPayMentType.ListIndex = 2
        End If
        
        txt_general_des = mgeneral_des
        '
        Set rsDummyUnit = New ADODB.Recordset
        s = "Select * from TblBranchesData Where Branch_Name Like '%" & Trim(mBranch_Name) & "%' Or  Branch_Code = '" & Trim(mBranch_Name) & "'"
                 
        rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummyUnit.EOF Then
            mbranch_id = val(rsDummyUnit!branch_id & "")
        Else
            mbranch_id = Current_branch
        End If
        rsDummyUnit.Close
    
        If CboPayMentType.ListIndex <> 2 Then
            s = " SELECT     BoxId,BoxName from tblBoxesData Where BoxName Like '%" & Trim(mBoxName) & "%'"
         
            rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsDummyUnit.EOF Then
                mBoxID = val(rsDummyUnit!BoxID & "")
            End If
         
            rsDummyUnit.Close
        End If
        If CboPayMentType.ListIndex <> 0 Then
            s = ""
            If mCode <> "" Then
                s = " SELECT     CusID,CusName from TblCustemers Where  FullCode = '" & Trim(mCode) & "'"
            Else
                s = " SELECT     CusID,CusName from TblCustemers Where CusName = N'%" & Trim(mCusName) & "%' "
        
            End If

            Set rsDummyUnit = New ADODB.Recordset
            rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly
            If rsDummyUnit.EOF Then
                rsDummyUnit.Close
        
                s = " SELECT     CusID,CusName from TblCustemers Where CusName Like N'%" & Trim(mCusName) & "%' "
                rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly
            End If
            If Not rsDummyUnit.EOF Then
                mCusId = val(rsDummyUnit!CusID & "")
            End If
    
            rsDummyUnit.Close
     
        End If
   
        s = " SELECT     Account_Code,Account_Name from Accounts Where Account_Name = '" & Trim(mAccountName) & "'"
    
        rsDummyUnit.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummyUnit.EOF Then
            mAccountCode = Trim(rsDummyUnit!Account_code & "")
        End If
    
        rsDummyUnit.Close
        mIsSaveFromExcel = True
    
        '    Set rsDummySupp = New ADODB.Recordset
        '    s = "SELECT CusID FROM TblCustemers Where (CusName Like '%" & Trim(mDefaultSupplier) & "%'     Or CusNamee Like '%" & Trim(mDefaultSupplier) & "%')"
        '    rsDummySupp.Open s, Cn, adOpenStatic, adLockReadOnly
        '    If Not rsDummySupp.EOF Then
        '        mDefaultSupplierID = val(rsDummySupp!CusID & "")
        '    End If
        '
        TxtModFlg.text = "N"
        isCmdCreateTransa = True
    
        XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
        '  Me.TxtNoteSerial.Text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
        Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=85"))
        Me.DCboUserName.BoundText = user_id
        '        XPDtbTrans.SetFocus
              
        XPDtbTrans.value = mNoteDate
        dcBranch.BoundText = mbranch_id
        If mCusId <> 0 Then
            DCVendor.BoundText = mCusId
            DCVendor_Click 0
        End If
        If mBoxID <> 0 Then
            DcboBox.BoundText = mBoxID
            DcboBox_Change
        End If
        dcBranch.BoundText = mbranch_id
        '               Dcbranch.BoundText = mbranch_id
        '               Dcbranch.BoundText = mbranch_id
        '
        Fg_Journal.rows = 1
        Fg_Journal.rows = 2
        Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("AccountName")) = mAccountName
        Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("AccountCode")) = mAccountCode
        Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("Count")) = 1
        Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("price")) = mPrice
        'Fg_Journal_AfterEdit 1, Fg_Journal.ColIndex("AccountName")
        Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("ExpensesID")) = get_Revenue_id(mAccountCode)
        Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("LineNo1")) = setfoxy_Line
        Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("Order_No")) = TXT_order_no.text
        Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("Count")) = 1
        Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("des")) = mgeneral_des
                
        'ddVATExp 2
            
    If ComResid(1).value Then
                    AddVATExp 1
                    AddVATExp 1
    End If
   

        ReLineGrid
        Fg_Journal.Row = 1
    
    If ComResid(1).value Then
                    AddVATExp 1
                    AddVATExp 1
    End If
   
                
        '     Fg_Journal_AfterEdit 2, Fg_Journal.ColIndex("price")
            
        setfoxy
        TxtSerial1.text = ""
                
        TxtSerial = ""
        TxtNoteSerial = ""
        TxtNoteSerial1 = ""
                
        SaveData "", ""
  
        If mBranch_Name = "" Then GoTo NextRow

NextRow:

    Next

    mIsSaveFromExcel = False
End Sub

Private Sub CmdPrintForms_Click(Index As Integer)
    SaveQRCode6 "notes_all", "NoteID", val(XPTxtID.text), TxtSerial1.text, (XPDtbTrans.value), _
       (XPTxtValView.text), Picture1, 0, (txtVatTotal.text), (XPTxtValView.text), val(dcBranch.BoundText)
    
    If DoPremis(Do_Print, Me.Name, True) = False Then
        Exit Sub
    End If

    print_report2 TxtSerial.text, DCVendor.text, Index

End Sub

Private Sub CmdRemove_Click()
    
    Dim X As Integer
    If val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("FlgVat"))) = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ Õ–› ”ÿ— «·›«  .Ì—ÃÏ  ’›Ì— ‰”»… «·›« "
        Else
            MsgBox "Can not delete VAT  "
        End If
        Exit Sub
    End If
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then
        Exit Sub
    End If
    Dim sql As String

    sql = "Delete  marakes_taklefa_temp where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
    Cn.Execute sql, , adExecuteNoRecords
    
    If CboPaymentType1.ListIndex = 0 Then
        If Fg_Journal.rows > 1 Then
            If Fg_Journal.rows = 2 Then
                Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.Fg_Journal.rows > 1 Then
                    If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                                        
                        With Me.Fg_Journal

                            '     If Me.TxtModFlg <> "E" Then Exit Sub
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
'
'                            LogTextA = "  Õ–› «·„’—Ê›   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " »ﬁÌ„… " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
'                            LogTexte = "  Delete  Expensen   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " With Value " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
'
'                            AddToLogFile CInt(user_id), 8063, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", val(Me.TxtSerial), val(TxtSerial1)
                        End With
                                                        
                        Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                    End If
                End If
            End If
        End If
            
        With Fg_Journal
            Me.XPTxtVal.text = Round(val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))), Decimal_Places)
            Me.XPTxtVal2.text = Round(val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value")) - .Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat"))), Decimal_Places)
        End With
        txtTotal = Round(val(Me.XPTxtVal.text), Decimal_Places)
        Me.XPTxtVal.text = Round(val(Me.XPTxtVal.text) - val(txtTotalFines), Decimal_Places)
    ElseIf CboPaymentType1.ListIndex = 1 Then

        If VSFlexGrid1.rows > 1 Then
            If VSFlexGrid1.rows = 2 Then
                Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.VSFlexGrid1.rows > 1 Then
                    If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                        
                        With Me.VSFlexGrid1

                            '   If Me.TxtModFlg <> "E" Then Exit Sub
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
                                                         
'                            LogTextA = "  Õ–› «·Õ”«»   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " »ﬁÌ„… " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
'                            LogTexte = "  Delete  Account   " & .Cell(flexcpTextDisplay, .Row, .ColIndex("AccountName")) & " With Value " & .Cell(flexcpTextDisplay, .Row, .ColIndex("Value"))
'
'                            AddToLogFile CInt(user_id), 8063, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", val(Me.TxtSerial), val(TxtSerial1)
                        End With
                        
                        Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                    End If
                End If
            End If
        End If
            
        With VSFlexGrid1
            Me.XPTxtVal.text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value")), Decimal_Places)
            Me.XPTxtVal2.text = Round(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value")) - .Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat")), Decimal_Places)
        End With
        txtTotal = Round(val(Me.XPTxtVal.text), Decimal_Places)
        Me.XPTxtVal.text = Round(val(Me.XPTxtVal.text) - val(txtTotalFines), Decimal_Places)
    Else
 
        Exit Sub
    End If

End Sub

Private Sub CMDSelectFile_Click()
    Cmd_Click 0
    On Error GoTo eh
    CommonDialog1.DialogTitle = "Select Upload list"
    CommonDialog1.CancelError = True
    CommonDialog1.filter = "xls Files (*.xls)|*.xls"
    CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNLongNames Or cdlOFNExplorer
    CommonDialog1.ShowOpen
    Dim vFiles() As String

    'grdFiles.Visible = True
    
    grdFiles.rows = 1
    Dim Row As Integer
    Row = 1
    Dim i As Integer
    vFiles = Split(CommonDialog1.FileName, CHR(0))
    If UBound(vFiles) = 0 Then
        grdFiles.AddItem Row
        grdFiles.TextMatrix(Row, grdFiles.ColIndex("File")) = Row
        grdFiles.TextMatrix(Row, grdFiles.ColIndex("Rows")) = 0
        grdFiles.TextMatrix(Row, grdFiles.ColIndex("DRows")) = 0
        grdFiles.TextMatrix(Row, grdFiles.ColIndex("FileName")) = CommonDialog1.FileName
    Else
        txtMainPath = vFiles(0)
        For i = 1 To UBound(vFiles)
            grdFiles.AddItem Row
            grdFiles.TextMatrix(Row, grdFiles.ColIndex("File")) = Row
            grdFiles.TextMatrix(Row, grdFiles.ColIndex("Rows")) = 0
            grdFiles.TextMatrix(Row, grdFiles.ColIndex("DRows")) = 0
            grdFiles.TextMatrix(Row, grdFiles.ColIndex("FileName")) = vFiles(0) & "\" & vFiles(i)
            Row = Row + 1
        Next
    End If
    txtFile.text = CommonDialog1.FileName
    LoadExcel
   
    Exit Sub
eh:
    MsgBox Err.Description
End Sub

Private Sub ComResid_Click(Index As Integer)
ClculteVAT
End Sub
Sub ClculteVAT()
If Me.TxtModFlg.text <> "R" Then
Dim Percetage As Double
Dim account2 As String
Dim account As String
Dim Percetage2 As Double
If ComResid(1).value = True Then
PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, account, Percetage

PercentgValueAddedAccount_Transec XPDtbTrans.value, 21, 1, account2, Percetage2

'TxtFATYou2.text = Percetage2
If SystemOptions.OpenVATAccountOwner = True Then
'TxtFATYou.text = 0

Else
'TxtFATYou.text = Percetage

End If
Else
'TxtFATYou.text = 0
'AccountVat.BoundText = ""
End If
'Calculte
End If
End Sub

Private Sub DcboBankName_Change()
    On Error Resume Next

    If DcboBankName.BoundText = "" Then
        Exit Sub
    End If
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

        '    Me.DcboCreditSide.BoundText = "a2a3a2"
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & val(DcboBankName.BoundText)
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
        If SystemOptions.banks_Accounts3 = True Then
            Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code2")
        Else
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
    
        'Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
  
        If CboPayMentType.ListIndex = 3 Then
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If

    End If

End Sub

Private Sub DcboBox_Change()

    If DcboBox.BoundText = "" Then
        Exit Sub
    End If
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub

Private Sub DcboBox_Click(Area As Integer)
    DcboBox_Change
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtSerial.text = ""
    TxtSerial1.text = ""
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 3
    End If

End Sub

Private Sub dcproject_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
             FrmProjectSearch.lblSearchtype.Caption = 666
             FrmProjectSearch.show vbModal
           
        End If
End Sub

Private Sub dcproject2_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
            FrmProjectSearch.lblSearchtype.Caption = 777
            FrmProjectSearch.show vbModal
           
        End If
End Sub

Private Sub DCVendor_Click(Area As Integer)
    Dim Balance       As String
    Dim balanceString As String
    If DCVendor.BoundText = "" Then
        Exit Sub
    End If
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DCVendor.BoundText))
    End If
 
    WriteCustomerBalPublic Me.DcboCreditSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
    If Me.TxtModFlg.text <> "R" Then
        '    TxtCurrentBalance.Text = Balance
    End If
   
End Sub
Sub DeleteGridCurrRowExp(Optional CurrRow As Long)
Dim i As Integer
With Fg_Journal
i = .rows
Do
i = i - 1
If val(.TextMatrix(i, .ColIndex("CurrRow"))) = CurrRow Then
.RemoveItem i
End If
Loop While i > 1
End With
End Sub
Sub HidFat()
    With Fg_Journal
        If True = True Then
            .ColHidden(.ColIndex("Vat")) = False
            .ColHidden(.ColIndex("Vatyo")) = False
        Else
            .ColHidden(.ColIndex("Vat")) = True
            .ColHidden(.ColIndex("Vatyo")) = True
        End If
    End With
End Sub
Sub AddVATExp(Optional Row As Long)
    If True = True Then
        Dim ForcedFlg      As Integer
        Dim valuee         As Double
        Dim AccountVATDept As String
        Dim i              As Integer
        Dim k              As Integer
        Dim ClsAcc         As New ClsAccounts
        Dim account        As String
        Dim Percetage      As Double
        Dim s              As String
        Dim rsDummy        As New ADODB.Recordset

        With Fg_Journal
            valuee = val(.TextMatrix(Row, .ColIndex("value")))

            .TextMatrix(Row, .ColIndex("Vatyo")) = PercentgValueAddedAccount(XPDtbTrans.value, .TextMatrix(Row, .ColIndex("AccountCode")), val(dcBranch.BoundText), ForcedFlg)
            If val(.TextMatrix(Row, .ColIndex("Vatyo"))) = 0 Then
                s = "Select * from TblSettsReqLimKDet Where  ID = -1"
                rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
                rsDummy.AddNew
                rsDummy("SetReqLID").value = 3
                rsDummy("Account_Code").value = .TextMatrix(Row, .ColIndex("AccountCode"))
                rsDummy("BranchID").value = IIf(val(dcBranch.BoundText) = 0, 1, val(dcBranch.BoundText))
                rsDummy("PercentD").value = 15
                rsDummy("ForcedFlg").value = False
                rsDummy("Typ").value = 9
                rsDummy.update
                .TextMatrix(Row, .ColIndex("Vatyo")) = PercentgValueAddedAccount(XPDtbTrans.value, .TextMatrix(Row, .ColIndex("AccountCode")), val(dcBranch.BoundText), ForcedFlg)
            End If
            If AkarPayCheck.value = vbChecked Then
                PercentgValueAddedAccount_Transec XPDtbTrans.value, 50, 1, account, Percetage
                .TextMatrix(Row, .ColIndex("Vatyo")) = Percetage
            End If

            If val(txtManulaVat.text) > 0 Then
                .TextMatrix(Row, .ColIndex("Vatyo")) = val(txtManulaVat.text)

            End If

            .TextMatrix(Row, .ColIndex("ForcedFlg")) = ForcedFlg
            .TextMatrix(Row, .ColIndex("Vat")) = Round((val(.TextMatrix(Row, .ColIndex("Vatyo"))) * valuee) / 100, Decimal_Places)
            .TextMatrix(Row, .ColIndex("DailyMonthly")) = ""
            GetValueAddedAccount XPDtbTrans.value, , AccountVATDept

            If AkarPayCheck.value = vbChecked Then
                AccountVATDept = account
 
            End If

            ''/////////////
            If val(.TextMatrix(Row, .ColIndex("Vat"))) > 0 Then
                If Not .TextMatrix(.Row, .ColIndex("AccountCode")) = "" Then
                    DeleteGridCurrRowExp Row
                    For i = 1 To 1
                        .AddItem " ", .Row + i
                        k = .Row + i
                        .TextMatrix(k, .ColIndex("CurrRow")) = Row
                        If i = 1 Then
                            '.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(AccountVATDept)
                            .TextMatrix(k, .ColIndex("AccountName")) = Get_Account_Name(, AccountVATDept)
                            .TextMatrix(k, .ColIndex("AccountCode")) = AccountVATDept
                        Else
                            .TextMatrix(k, .ColIndex("AccountCode")) = DcboCreditSide.BoundText
                            .TextMatrix(k, .ColIndex("AccountName")) = Get_Account_Name(, DcboCreditSide.BoundText)
                            '.TextMatrix(k, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(DcboCreditSide.BoundText)
                        End If
                        .TextMatrix(k, .ColIndex("price")) = .TextMatrix(Row, .ColIndex("Vat"))
                        .TextMatrix(k, .ColIndex("value")) = .TextMatrix(Row, .ColIndex("Vat"))
                        .TextMatrix(k, .ColIndex("Count")) = 1
                        .TextMatrix(k, .ColIndex("ExpensesID")) = .TextMatrix(Row, .ColIndex("ExpensesID"))
                        If SystemOptions.UserInterface = ArabicInterface Then
                            .TextMatrix(k, .ColIndex("des")) = .TextMatrix(Row, .ColIndex("des")) & " " & " ﬁÌ„… „÷«›…"
                        Else
                            .TextMatrix(k, .ColIndex("des")) = .TextMatrix(Row, .ColIndex("des")) & " " & " VAT"
                        End If
                        .TextMatrix(k, .ColIndex("FlgVat")) = 1
                        .TextMatrix(k, .ColIndex("CarId")) = .TextMatrix(Row, .ColIndex("CarId"))
                        .TextMatrix(k, .ColIndex("Order_No")) = .TextMatrix(Row, .ColIndex("Order_No"))
                        .TextMatrix(k, .ColIndex("CarName")) = .TextMatrix(Row, .ColIndex("CarName"))
                        .TextMatrix(k, .ColIndex("opr_fullcode")) = .TextMatrix(Row, .ColIndex("opr_fullcode"))
                    Next i
                End If
            End If
        End With
    End If
End Sub

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
   ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg            As String
    Dim rs             As New ADODB.Recordset
    Dim StrSQL         As String
    Dim ClsAcc         As New ClsAccounts
    Dim LngRow         As Long

    With Fg_Journal
        If .TextMatrix(Row, .ColIndex("CarName")) = "" Then
            .TextMatrix(Row, .ColIndex("CarName")) = DcbCar.text
            .TextMatrix(Row, .ColIndex("CarID")) = val(DcbCar.BoundText)
        End If

        Select Case .ColKey(Col)

            Case "CarName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
     
                StrAccountCode = .ComboData
            
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CarID"), False, True)
                .TextMatrix(Row, .ColIndex("CarID")) = StrAccountCode
            
                .TextMatrix(Row, .ColIndex("des")) = "’—›  ⁄·Ï «·„⁄œÂ/«·”Ì«—…  : " & .TextMatrix(Row, .ColIndex("CarName"))
            Case "Vatyo"
                If val(.TextMatrix(Row, .ColIndex("Vatyo"))) = 0 Then
                    .TextMatrix(Row, .ColIndex("Vat")) = 0
                    If .rows > Row Then
                        If val(.TextMatrix(Row + 1, .ColIndex("FlgVat"))) = 1 Then
                            .RemoveItem Row + 1
                        End If
                    End If
                End If
       
            Case "ExpensesID"
              
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
 
            Case "FromCityName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CityFromId"), False, True)
                .TextMatrix(Row, .ColIndex("CityFromId")) = StrAccountCode
 
            Case "ToCityName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CityToId"), False, True)
                .TextMatrix(Row, .ColIndex("CityToId")) = StrAccountCode
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Revenue_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
                .TextMatrix(Row, .ColIndex("Order_No")) = TXT_order_no.text
                .TextMatrix(Row, .ColIndex("Count")) = 1
                If ComResid(1).value Then
                    AddVATExp Row
                End If
                ' If SystemOptions.UserInterface = ArabicInterface Then
                '     StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                '
                ' Else
                '     StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                ' End If
            
                ' rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                ' If rs.RecordCount > 0 Then
                '     .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                ' Else
                '     .TextMatrix(Row, .ColIndex("des")) = ""
                ' End If
            Case "Count", "price", "Discount"
                ReLineGrid
                If ComResid(1).value Then
                    AddVATExp Row
                End If
                
            Case "value", "opr_fullcode"
                Dim sgl        As String
                Dim project_id As Integer
                project_id = get_project_id(dcproject.BoundText, "REVENUE_account")
                
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                End If
                If ComResid(1).value Then
                    AddVATExp Row
                End If
                Me.XPTxtVal.text = Round(val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value")) - val(txtPayAmount)), Decimal_Places)
                txtTotal = Round(val(Me.XPTxtVal.text), Decimal_Places)
                Me.XPTxtVal.text = Round(val(Me.XPTxtVal.text) - val(txtTotalFines), Decimal_Places)
                
                Me.XPTxtVal2.text = Round(val(Me.XPTxtVal.text) - .Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat")), Decimal_Places)
                sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.XPTxtVal.text = Round(val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))) - val(txtPayAmount), Decimal_Places)
        txtTotal = Round(val(Me.XPTxtVal.text), Decimal_Places)
        Me.XPTxtVal.text = Round(val(Me.XPTxtVal.text) - val(txtTotalFines), Decimal_Places)
                
        Me.XPTxtVal2.text = Round(val(Me.XPTxtVal.text) - .Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat")), Decimal_Places)
        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid

    With Me.Fg_Journal

        If Me.TxtModFlg <> "E" Then
            Exit Sub
        End If

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
'        If Col = .ColIndex("Account_Serial") Or Col = .ColIndex("AccountName") Then
'            LogTextA = "   ⁄œÌ· «·„’—Ê› «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'            LogTexte = "  Change Account To " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'        ElseIf Col = .ColIndex("Value") Then
'            LogTextA = "   ⁄œÌ· «·ﬁÌ„…  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " ··„’—Ê›   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'            LogTexte = "  Change value" & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " To Expenses " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'        ElseIf Col = .ColIndex("Des") Then
'            LogTextA = "   ⁄œÌ· «·‘—Õ  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " ··„’—Ê›   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'            LogTexte = "  Change Des " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " To Expenses " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'        End If
'
'        AddToLogFile CInt(user_id), 8063, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
    
    End With

End Sub

Function calcnets()

    If Me.CboPaymentType1.ListIndex = 0 Then

        With Fg_Journal
                Me.XPTxtVal.text = Round(val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))) - val(txtPayAmount), Decimal_Places)
                txtTotal = Round(val(Me.XPTxtVal.text), Decimal_Places)
                Me.XPTxtVal.text = Round(val(Me.XPTxtVal.text) - val(txtTotalFines), Decimal_Places)
                
                Me.XPTxtVal2.text = Round(val(Me.XPTxtVal.text) - .Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat")), Decimal_Places)
            txtVatTotal = Round(val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat"))), Decimal_Places)
        End With

    Else

     '   With Me.VSFlexGrid1

'      Me.XPTxtVal.text = val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))) - val(txtPayAmount)
'                txtTotal = val(Me.XPTxtVal.text)
'                Me.XPTxtVal.text = val(Me.XPTxtVal.text) - val(txtTotalFines)
'                If VSFlexGrid1.rows > VSFlexGrid1.FixedRows Then
'                    If val(.TextMatrix(1, .ColIndex("Vat"))) <> 0 Then
'                Me.XPTxtVal2.text = val(Me.XPTxtVal.text) - val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat")))
'
'            txtVatTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat"))
'            End If
'            End If
'        End With

    End If

End Function

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

     If val(.TextMatrix(Row, .ColIndex("FlgVat"))) <> 0 Then
            Cancel = True
     Else
 Select Case .ColKey(Col)
        Case "Vat"
                 Cancel = True
        Case "Vatyo"
              If val(.TextMatrix(Row, .ColIndex("ForcedFlg"))) = 1 Then
                 Cancel = True
              Else
              .ComboList = ""
              End If

            Case "value"
              Cancel = True
                .ComboList = ""
                
            Case "Count"
              .ComboList = ""
              
              Case "price"
              .ComboList = ""
              
            Case "Discount"
              .ComboList = ""
            Case "MonthCount", "PurchOrderNo", "DailyMonthlyValue"
                .ComboList = ""
            Case "des"
                .ComboList = ""
                '  Cancel = True
               Case "des2"
                .ComboList = ""
            Case "Order_No"
                .ComboList = ""
        End Select
   End If
  End With

End Sub

Private Sub Fg_Journal_DblClick()
    Exit Sub
  
    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" Then
        
            CboDes.Visible = False
       
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
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
        If r <= 0 Or c <= 0 Then
            Exit Sub
        End If
        If typename(Fg_Journal.cell(flexcpData, r, c)) <> "String" Then
            TxtDes.text = ""
        Else
            '
            TxtDes.text = Fg_Journal.cell(flexcpData, r, c)
        End If

        ' show new note
        

        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus

        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyPress(KeyAscii As Integer)
 Sendkeys "{F4}"
  Sendkeys "{BACKSPACE}"
  Sendkeys CHR(KeyAscii)
End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    With Fg_Journal

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
                 
                    Order_no_search.show
                    Order_no_search.RetrunType = 4

                End If

            Case "AccountName"

                If KeyCode = vbKeyF3 Then
                    FrmExpensesSearch.RetrunType = 23
                    FrmExpensesSearch.Indx = 3
                    FrmExpensesSearch.show
                    

                End If
            
 
        End Select

    End With

End Sub

Public Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String
    Dim StrComboList2 As String
    Dim StrComboListCity As String
    
    
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)
                
                '«ŸÂ«— «·„⁄œ« /«·”Ì«—« 
            Case "CarName"
        
                StrSQL = "  select id,BoardNO from TblCarsData"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList2 = Fg_Journal.BuildComboList(rs, "BoardNO", "id")
       
                If StrComboList2 <> "" Then
                    StrComboList2 = "|" & StrComboList2
                End If
                .ColComboList(Col) = StrComboList2
                '.ComboList = StrComboList2
          Case "FromCityName", "ToCityName"
                If .ColKey(Col) = "FromCityName" Then
                             StrSQL = " SELECT     dbo.TblCountriesGovernments.GovernmentID CityID, dbo.TblCountriesGovernments.GovernmentName "
                    StrSQL = StrSQL & " FROM         dbo.TblCountriesGovernments left OUTER JOIN"
                    StrSQL = StrSQL & " dbo.TBLCitiesDistance ON dbo.TblCountriesGovernments.GovernmentID = dbo.TBLCitiesDistance.CityFromId"
                    StrSQL = StrSQL & "  GROUP BY dbo.TblCountriesGovernments.GovernmentName, dbo.TblCountriesGovernments.GovernmentID"
                Else
                             StrSQL = "  SELECT      dbo.TblCountriesGovernments.GovernmentID CityID, dbo.TblCountriesGovernments.GovernmentName"
                    StrSQL = StrSQL & " FROM         dbo.TblCountriesGovernments left OUTER JOIN"
                    StrSQL = StrSQL & " dbo.TBLCitiesDistance ON dbo.TblCountriesGovernments.GovernmentID = dbo.TBLCitiesDistance.CityToId"
                    StrSQL = StrSQL & " GROUP BY dbo.TblCountriesGovernments.GovernmentName, dbo.TblCountriesGovernments.GovernmentID"
                End If
                StrSQL = StrSQL + "  ORDER BY dbo.TblCountriesGovernments.GovernmentName"
                

                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboListCity = Fg_Journal.BuildComboList(rs, "GovernmentName", "CityID")
       
                If StrComboListCity <> "" Then
                    StrComboListCity = "|" & StrComboListCity
                End If

                .ColComboList(Col) = StrComboListCity

            Case "AccountName"
                 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "SELECT     ACCOUNTS_1.Account_Code, ACCOUNTS_1.Account_Name FROM         dbo.ACCOUNTS ACCOUNTS_1 RIGHT OUTER JOIN dbo.TblRevenuesTypes ON ACCOUNTS_1.Account_Code = dbo.TblRevenuesTypes.Account_Code  order by Account_Name"
                Else
                    StrSQL = "SELECT     ACCOUNTS_1.Account_Code, ACCOUNTS_1.Account_NameEng FROM         dbo.ACCOUNTS ACCOUNTS_1 RIGHT OUTER JOIN dbo.TblRevenuesTypes ON ACCOUNTS_1.Account_Code = dbo.TblRevenuesTypes.Account_Code  order by Account_NameEng"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ColComboList(Col) = StrComboList
                  
            Case "opr_fullcode"
                Dim project_id As Integer
                project_id = get_project_id(dcproject.BoundText, "REVENUE_account")

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

                .ColComboList(Col) = StrComboList1
         
        End Select

    End With

End Sub

Private Sub Form_Load()
   
    Dim StrSQL As String
    'On Error GoTo ErrTrap
    
    If SystemOptions.IsCahngeServiceInvoice Then
        ScreenNameArabic = "›« Ê—… ÷—Ì»Ì… - «œ«—… «·«„·«ﬂ"
        ScreenNameEnglish = "Tax Invoice"

    Else
        ScreenNameArabic = "›« Ê—… Œœ„Ì…"
        ScreenNameEnglish = "Financial Invoice"
    End If
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 8063
    If SystemOptions.UserInterface = EnglishInterface Then
        Me.Caption = ScreenNameEnglish
        Ele.Caption = ScreenNameEnglish
    Else
        Me.Caption = ScreenNameArabic
        Ele.Caption = ScreenNameArabic

    End If
    If SystemOptions.AllowEditVaTManulay = True Then
        txtManulaVat.Enabled = True
        txtManulaVat.Visible = True
    Else
        txtManulaVat.Enabled = False
        txtManulaVat.text = 0
        txtManulaVat.Visible = False
    End If
    txtFromDate.value = Date
    txtToDate.value = Date
    txtFromDateH.value = ToHijriDate(txtFromDate.value)
    txtToDateH.value = ToHijriDate(txtToDate.value)

    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL
    HidFat
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
    
           With Me.DefaultInvoicetype
            .Clear
            
             


            .AddItem " ›« Ê—… ÷—Ì»Ì…  "
            .ItemData(0) = 0
     
            .AddItem " ›« Ê—… ÷—Ì»Ì… „»”ÿ… "
            .ItemData(1) = 2
         
        End With

    
    Set Dcombos = New ClsDataCombos
    Dcombos.GetDocTypebyid Me.DCDocTypes, 21, val(Me.dcBranch.BoundText)
    
    
     StrSQL = " select id,code from currency"
 
    fill_combo Me.Dccurrency, StrSQL
    
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetExpensesType XPCboExpensesType
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetCars Me.DcbCar
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.XPCboExpensesType

    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide

    With Me.CboPayMentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "«Ã·"
        .AddItem "‘Ìﬂ „”œœ"
        .AddItem "„œÏ"
        .AddItem "›Ì“«"
        '.AddItem "Õ”«»  "
    End With

    With Me.CboPaymentType1
        .Clear
        .AddItem "Œœ„…"
        .AddItem "Õ”«»« "
    
    End With

    With Me.CBoBasedON
        .Clear
        .AddItem "»·«"
        .AddItem "√„— ‘—¡"
        .AddItem "›« Ê—… „»œ∆ÌÂ"
        .AddItem " «„— «‰ «Ã  "
    
    End With
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        Fg_Journal.ColComboList(Fg_Journal.ColIndex("DailyMonthly")) = "#0;Daily|#1;Monthly"
    Else
    
        Fg_Journal.ColComboList(Fg_Journal.ColIndex("DailyMonthly")) = "#0;ÌÊ„Ì|#1;‘Â—Ì"
    End If
   
    StrSQL = " select REVENUE_account,Project_name from projects  where not(REVENUE_account is null) order by Project_name"
    fill_combo dcproject, StrSQL

    'StrSQL = " select  CusID, CusName from TblCustemers  where Type=2"
    'fill_combo Me.DCVendor, StrSQL

    Dcombos.GetCustomersSuppliers 1, Me.DCVendor

    Set rs = New ADODB.Recordset
'    StrSQL = "select * From notes_all where notetype=85 and bill_Type<>2"
'    If SystemOptions.usertype <> UserAdminAll Then
'        StrSQL = StrSQL & " AND   branch_no=" & Current_branch
'    End If
    
'    StrSQL = "select  * From notes_all where notetype=85 and bill_Type <>2 AND branch_no in(" & Current_branchSql & ")"
'    If SystemOptions.FixedCustomer = 1 Then
'        StrSQL = StrSQL & " and  UserID = " & user_id
'    End If
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Sub MoveTo(nav As Integer, Optional ByVal mTransID As Long = 0)
TryAgain:
    Dim s As String
    s = "Select Top 1 * "
    s = s & " from notes_all where notetype=85 and bill_Type <>2 "
    s = s & "  AND branch_no in ( " & Current_branchSql & ")"
    If SystemOptions.FixedCustomer = 1 Then
        s = s & " and  UserID = " & user_id
    End If
    Select Case nav
        Case TransNaV.EnFirstTrans
            s = s & " Order By NoteID  "
        Case TransNaV.EnLastTrans
            s = s & " Order By NoteID Desc"
        Case TransNaV.EnNextTrans
            If rs.EOF Or rs.BOF Then
                s = s & " Order By NoteID Desc"
            Else
                s = s & " And  NoteID  >" & rs!NoteID
            End If
           
        Case TransNaV.EnPervTrans
            If rs.EOF Or rs.BOF Then
                s = s & " Order By NoteID"
            Else
                s = s & " And NoteID < " & rs!NoteID
                s = s & "  Order By NoteID  DESC"
            End If
        Case TransNaV.EnSrchTrans
            
                s = s & " And NoteID = " & mTransID
                s = s & "  Order By NoteID  DESC"
           
        
           
    End Select
    Set rs = New ADODB.Recordset
    rs.Open s, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF And nav = EnPervTrans Then
        nav = EnFirstTrans
        GoTo TryAgain
    End If
    If rs.EOF And nav = EnNextTrans Then
        nav = EnLastTrans
        GoTo TryAgain
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 8063
    hide_logo = False

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

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, _
                               ByVal SpinningEnded As Boolean)

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
            TxtDes.text = Fg_Journal.cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Des"))
            CboDes.DropDown PicDes.hWnd, vdsRightToLeft, vdsBottomToDown, vdsDownArrow, True, vdsSoftResize
            Debug.Print PicDes.Height & "Pic H " & "-----" & PicDes.Width & "Pic W"
        Else
            CboDes.CloseUp
        End If
    End If

End Sub

Private Sub CboDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Sendkeys "{F4}"
    End If

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption



End Sub

 
 Private Sub LblLink_Click()
 
    Dim FirstPeriod As Date
    getFirstPeriodDateInthisYear FirstPeriod
    ShowReport DcboCreditSide.BoundText, DcboCreditSide.text, FirstPeriod, Date

End Sub



Private Sub PicDes_Resize()

    With PicDes
        LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
        TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
        '    PicHeight = PicDes.Height
        '    PicWidth = PicDes.Width
    End With

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
     '   GetCustomersDetail CUSTID, , Text2.Text, 1
        GetTblCustemersCode Text2.text, CUSTID
        DCVendor.BoundText = CUSTID
    End If

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 2
        FrmCustemerSearch.show vbModal
     
    End If
End Sub



Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If CBoBasedON.ListIndex = 3 Then
        If KeyCode = vbKeyF3 Then
            Order_no_search2.show
            Order_no_search2.RetrunType = 2
         
        End If

    Else

        If KeyCode = vbKeyF3 Then
            Order_no_search.show
            Order_no_search.RetrunType = 1
        End If

    End If

End Sub

Private Sub TxtDes_LostFocus()
    PicHeight = PicDes.Height
    PicWidth = PicDes.Width
    CboDes.CloseUp
    CboDes.Visible = False
End Sub

Private Sub TxtDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyEscape Then
        PutData
        CboDes.CloseUp
    End If

End Sub

Private Sub txtFromDate_Change()
If Me.TxtModFlg.text <> "R" Then

             txtFromDateH.value = ToHijriDate(txtFromDate.value)

End If

End Sub

Private Sub txtFromDateH_LostFocus()
    If Me.TxtModFlg.text <> "R" Then

        VBA.Calendar = vbCalGreg
        txtFromDate.value = ToGregorianDate(txtFromDateH.value)
           
    End If
End Sub

Private Sub txtManulaVat_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "N" Then
        
        Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.rows = 3
            Fg_Journal.Enabled = True
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
        
            XPTxtVal.locked = True
            '        XPCboProfLevel.Locked = True
            '        XPTxtProfMail.Locked = True
            '        XPTxtPhone.Locked = True
            '        XPTxtMobile.Locked = True
            XPMTxtRemarks.locked = True
            XPCboExpensesType.locked = True
            Me.DcboBox.locked = True
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
        
            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
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
        
            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtNoteSerial1_Change()
If TxtNoteSerial1.text <> "" Then
Dim Type1 As Integer
Dim txtperson As String
Dim des As String
Dim EmpID As Integer
Dim Price As Double
If Me.TxtModFlg.text <> "R" Then
OrderExchange TxtNoteSerial1.text, Type1, txtperson, des, Price, EmpID
CboPayMentType.ListIndex = Type1
'txtto.text = txtperson
txt_general_des.text = des
End If
End If
End Sub

Private Sub TxtNoteserial1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
       
            FrmReqExchangeSearch.show
            FrmReqExchangeSearch.lbltype.Caption = 3
          
        End If
End Sub

Private Sub txtPayAmount_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    calcnets
End If
End Sub

Private Sub txtToDate_Change()
If Me.TxtModFlg.text <> "R" Then

             txtToDateH.value = ToHijriDate(txtToDate.value)

End If

End Sub

Private Sub txtToDateH_LostFocus()
If Me.TxtModFlg.text <> "R" Then

           VBA.Calendar = vbCalGreg
           txtToDate.value = ToGregorianDate(txtToDateH.value)
           
End If
End Sub

Private Sub txtTotalFines_Change()
    calcnets
End Sub

Public Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
   ByVal Col As Long)
    'check_cost_center
    Dim StrAccountCode As String
    Dim Msg            As String
    Dim rs             As New ADODB.Recordset
    Dim StrSQL         As String
    Dim ClsAcc         As New ClsAccounts
    Dim LngRow         As Long
    Dim sql            As String
 
    With VSFlexGrid1

        Select Case .ColKey(Col)
    
            Case "Value"
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
'                Me.XPTxtVal.text = Round(val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))) - val(txtPayAmount), Decimal_Places)
'                txtTotal = Round(val(Me.XPTxtVal.text), Decimal_Places)
'                Me.XPTxtVal.text = Round(val(Me.XPTxtVal.text) - val(txtTotalFines), Decimal_Places)
                
              '  Me.XPTxtVal2.text = Round(val(Me.XPTxtVal.text) - .Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat")), Decimal_Places)

            Case "DebitValue", "CreditValue"

                'remove destribution
     
                ' sgl = "update  marakes_taklefa_temp  set value=0 where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                ' Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))
            
                If .ColKey(Col) = "DebitValue" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                    ' Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                 
                    '    Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValue" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                    ' Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '     Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If

                .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
            
            Case "DebitValueE", "CreditValueE"
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))

                If .ColKey(Col) = "DebitValueE" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
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
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
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

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
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

                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    
                    End If
                    
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)
                    
                    Dim rs2    As ADODB.Recordset
                    Dim My_SQL As String

                    If IsNull(rs("currenct_code").value) Then

                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                    
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo xx
                    End If

                    My_SQL = "  select * from currency WHERE id=" & val(rs("currenct_code").value)

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
xx:
                Else
                    'GetMsgs 130, vbExclamation
                    MsgBox "ﬂÊœ Õ”«» Œ«ÿÏ¡", vbCritical
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                
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
            
                StrSQL = "SELECT ACCOUNTS.cost_center ,ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Name='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), vbFalse, rs("cost_center").value)
            
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

                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value)
ll:
                End If

        End Select

        'to Add new row if needed
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If

        ReLineGrid

    End With

    With Me.VSFlexGrid1

        If Me.TxtModFlg <> "E" Then
            Exit Sub
        End If

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
'        If Col = .ColIndex("Account_Serial") Or Col = .ColIndex("AccountName") Then
'            LogTextA = "   ⁄œÌ· «·Õ”«» «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'            LogTexte = "  Change Account To " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'        ElseIf Col = .ColIndex("Value") Then
'            LogTextA = "   ⁄œÌ· «·ﬁÌ„…  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'            LogTexte = "  Change value" & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'        ElseIf Col = .ColIndex("Des") Then
'            LogTextA = "   ⁄œÌ· «·‘—Õ  «·Ï " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " ··Õ”«»   " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'            LogTexte = "  Change Des " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " To Account " & .Cell(flexcpTextDisplay, Row, .ColIndex("AccountName"))
'        End If
'
'        AddToLogFile CInt(user_id), 8063, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", "", Me.TxtSerial, TxtSerial1
    End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

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

Private Sub VSFlexGrid1_KeyPress(KeyAscii As Integer)
Sendkeys "{F4}"
Sendkeys "{BACKSPACE}"
Sendkeys CHR(KeyAscii)
End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 8063

    End If

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

            Case "AccountName"
                
                'Full Path Display
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

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
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

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

Private Sub XPBtnMove_Click(Index As Integer)
    '    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If
    MoveTo Index
     Retrive
    '    Select Case Index
    '
    '        Case 0
    '
    '            '            If Not (rs.EOF Or rs.BOF) Then
    '            '                rs.MovePrevious
    '            '
    '            '                If rs.BOF Then rs.MoveFirst
    '            '            End If
    '
    '            MoveTo EnPervTrans
    '
    '        Case 1
    '
    ''            If Not (rs.EOF Or rs.BOF) Then
    ''                rs.MoveFirst
    ''            End If
    ' MoveTo EnFirstTrans
    '
    '        Case 2
    '
    ''            If Not (rs.EOF Or rs.BOF) Then
    ''                rs.MoveLast
    ''            End If
    ' MoveTo EnLastTrans
    '
    '        Case 3
    '
    ''            If Not (rs.EOF Or rs.BOF) Then
    ''                rs.MoveNext
    ''
    ''                If rs.EOF Then rs.MoveLast
    ''            End If
    ' MoveTo EnNextTrans
    '    End Select
' Retrive
   
    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim CarID As Integer
    Dim CarName As String
    If Lngid <> 0 Then MoveTo 4, Lngid
 '   On Error GoTo ErrTrap
    Fg_Journal.Clear flexClearScrollable, flexClearEverything
    Fg_Journal.rows = 3
          
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.rows = 2
          
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        'Lngid
        '  If XPTxtID.text <> 0 Then
        '      Rs.find "NoteID=" & XPTxtID.text, , adSearchForward, adBookmarkFirst
        '      If Rs.EOF Or Rs.BOF Then
        '          Exit Sub
        '      End If
        '  End If
  
        If Lngid <> 0 Then
            rs.Find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If

    End If

    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    Else
        Me.DcCostCenter.BoundText = ""
    End If
    ComResid(1).value = True
    Me.DcbCar.BoundText = IIf(IsNull(rs("CarID").value), "", rs("CarID").value)
    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.Text1.text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    Me.TXT_order_no.text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
    Me.txtOrderID.text = IIf(IsNull(rs("OrderID").value), "", rs("OrderID").value)
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("Noteseril2").value), "", rs("Noteseril2").value)
    TXT_A_NoteID.text = IIf(IsNull(rs("A_NoteID").value), "", val(rs("A_NoteID").value))
    txtPayAmount.text = rs("PayAmount") & ""

    XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
   ' XPTxtVal2.Text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
    
     txtVendor.text = IIf(IsNull(rs("Vendor").value), "", rs("Vendor").value)
     
      DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)
    Me.Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
    txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    txtDateRec.value = IIf(IsNull(rs("DateRec").value), Date, (rs("DateRec").value))
    zatcaStatus = IIf(IsNull(rs("zatcaStatus").value), 0, rs("zatcaStatus").value)
    TXTIban.text = IIf(IsNull(rs("CIBAN").value), "", (rs("CIBAN").value))
    
    
    If Not IsNull(rs.Fields("ComResid").value) Then
   If rs.Fields("ComResid").value = 1 Then
   ComResid(1).value = True
   Else
   ComResid(0).value = True
   End If
   Else
   ComResid(1).value = True
   End If
   TXTNewNO.text = IIf(IsNull(rs("NewNO").value), "", rs("NewNO").value)
    

    DefaultInvoicetype.ListIndex = IIf(IsNull(rs("Invoicetype").value), 0, rs("Invoicetype").value)
    
     Dim mmm As String
    
    If Not (IsNull(rs("QrCodeImage").value)) Then
        LoadPictureFromDB Picture2, rs, "QrCodeImage", mmm
    Else
     Set Picture2.Picture = Nothing
    End If

            If rs("AkarPayCheck").value = 1 Then
            AkarPayCheck.value = vbChecked

        Else
            AkarPayCheck.value = vbUnchecked
        End If


    XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    txtto.text = IIf(IsNull(rs("too").value), "", rs("too").value)
    txt_general_des.text = IIf(IsNull(rs("general_des").value), "", rs("general_des").value)
txtManulaVat.text = IIf(IsNull(rs("txtManulaVat").value), 0, (rs("txtManulaVat").value))
    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    XPCboExpensesType.BoundText = IIf(IsNull(rs("ExpensesID").value), "", rs("ExpensesID").value)

    txtFromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    
    txtFromDateH.value = IIf(IsNull(rs("FromDateH").value), "", rs("FromDateH").value)
    

    txtToDate.value = IIf(IsNull(rs("ToDate").value), Date, rs("ToDate").value)
    
    txtToDateH.value = IIf(IsNull(rs("ToDateH").value), "", rs("ToDateH").value)
    
    If (rs("bill_Type").value) = 0 Then
        Me.CboPaymentType1.ListIndex = 0
    ElseIf (rs("bill_Type").value) = 1 Then
        Me.CboPaymentType1.ListIndex = 1
'    ElseIf (rs("bill_Type").value) = 4 Then
'        Me.CboPaymentType1.ListIndex = 4
'    ElseIf (rs("bill_Type").value) = 5 Then
'        Me.CboPaymentType1.ListIndex = 5
    Else
        Me.CboPaymentType1.ListIndex = 0
    End If

    CboPaymentType1_Change

    If Not IsNull(rs("BasedONID").value) Then
        Me.CBoBasedON.ListIndex = rs("BasedONID").value
    Else
        Me.CBoBasedON.ListIndex = 0
 
    End If
Me.DtpChequeDueDate.value = IIf(IsNull(rs("DueDate").value), Date, rs("DueDate").value)
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

    ElseIf rs("NoteCashingType").value = 3 Then
        Me.CboPayMentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    
    ElseIf rs("NoteCashingType").value = 2 Then
        Me.CboPayMentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
    
        Me.DCVendor.BoundText = rs("CusID").value

    ElseIf rs("NoteCashingType").value = 4 Then
        Me.CboPayMentType.ListIndex = 4
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value

    ElseIf rs("NoteCashingType").value = 5 Then
        Me.CboPayMentType.ListIndex = 5
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value

    End If
    
   

    CboPayMentType_Change

    'ÿMe.DcboBox.BoundText = IIf(IsNull(Rs("BoxID").value), "", Rs("BoxID").value)
    'DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))

    If rs("NoteCashingType").value = 0 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    ElseIf rs("NoteCashingType").value = 1 Or rs("NoteCashingType").value = 3 Or rs("NoteCashingType").value = 4 Or rs("NoteCashingType").value = 5 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DcboBankName.BoundText))
    ElseIf rs("NoteCashingType").value = 2 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DCVendor.BoundText))
    End If

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.Txt_Numorder.text = IIf(IsNull(rs("NumOrderInpot").value), "", rs("NumOrderInpot").value)
    Me.TxtSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)

    Me.TxtSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    Me.txtTotalFines.text = IIf(IsNull(rs("TotalFines").value), "", rs("TotalFines").value)
    Me.txtRequestNo.text = IIf(IsNull(rs("RequestNo").value), "", rs("RequestNo").value)
    Me.txtContractNo.text = IIf(IsNull(rs("ContractNo").value), "", rs("ContractNo").value)

    Me.oldTxtSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(27).Caption = showLabel(TxtSerial1, oldTxtSerial1)

    Me.dcproject.BoundText = IIf(IsNull(rs("project_Expensen_account").value), "", rs("project_Expensen_account").value)

   ' If CboPaymentType1.ListIndex = 1 Then 'Õ”«Ì« 

        StrSQL = "SELECT     TOP 100 PERCENT   DEV_ID_Line_No1,dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, "
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.UserID , dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[value],dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description"
        StrSQL = StrSQL + " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
        StrSQL = StrSQL + " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
        StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1)  and isNull(DOUBLE_ENTREY_VOUCHERS.IsExpens,0) = 1 And  (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & val(Me.XPTxtID.text) & ")"
        StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"

        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If RsDev.RecordCount > 0 Then
            RsDev.MoveFirst
        End If
    
        With Me.VSFlexGrid1
 
            .rows = .FixedRows + RsDev.RecordCount
 
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                 .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), "", RsDev("DEV_ID_Line_No1").value)
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
            
                .TextMatrix(i, .ColIndex("account_serial")) = IIf(IsNull(RsDev("account_serial").value), "", RsDev("account_serial").value)
            
                If SystemOptions.UserInterface = ArabicInterface Then
            
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                Else
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                End If
                   
             
                
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
            
                .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), "", RsDev("Double_Entry_Vouchers_Description").value)
            
                RsDev.MoveNext
            Next i
    
        End With
       ' XPTxtVal_Change
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
      '  ReLineGrid
     '   Exit Sub
    'End If


Dim s As String

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then '«·„—Ê›« 
        '   StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(Me.XPTxtID.text)
        '   StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        ' StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.*,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name FROM    dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code WHERE     dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID =" & Val(Me.XPTxtID.text) & "Order By DEV_ID_Line_No"

        'StrSQL = "SELECT   dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode,   dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit , dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID ,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description  FROM         dbo.ACCOUNTS INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
        'StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0  and dbo.DOUBLE_ENTREY_VOUCHERS.notes_all =" & Val(Me.XPTxtID.text) & ") "
        'StrSQL = StrSQL + "ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
        StrSQL = "SELECT  dbo.DOUBLE_ENTREY_VOUCHERS.CarID , dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID, dbo.Notes.DailyMonthlyValue,"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Nameeng,"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID AS Expr1,"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description , dbo.Notes.order_no"
        StrSQL = StrSQL + "  , dbo.Notes.[Count], dbo.Notes.price  ,dbo.Notes.Discount,dbo.DOUBLE_ENTREY_VOUCHERS.CurrRow,dbo.DOUBLE_ENTREY_VOUCHERS.FlgVat,dbo.DOUBLE_ENTREY_VOUCHERS.Vatyo,dbo.DOUBLE_ENTREY_VOUCHERS.Vat,Notes.DailyMonthly "
        StrSQL = StrSQL + "  ,Notes.MonthCount,Notes.PurchOrderNo,Notes.CityFromId,Notes.CityToId,Notes.des2,Notes.ContItem,Notes.PurchaseOrderNo,Notes.LocationName,Notes.RentType,Notes.UnitNo"
        StrSQL = StrSQL + "  ,FromCity.GovernmentName FromCityName,ToCity.GovernmentName ToCityName"
        StrSQL = StrSQL + "  FROM         dbo.ACCOUNTS INNER JOIN"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
        StrSQL = StrSQL + " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
        StrSQL = StrSQL + "             LEFT OUTER JOIN TblCountriesGovernments FromCity ON Notes.CityFromId = FromCity.GovernmentID"
        StrSQL = StrSQL + "             LEFT OUTER JOIN TblCountriesGovernments ToCity ON Notes.CityToId = ToCity.GovernmentID"

        StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1)  and isNull(DOUBLE_ENTREY_VOUCHERS.IsExpens,0) = 0 And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & val(Me.XPTxtID.text) & ")"
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
    
            With Me.Fg_Journal

                If Me.dcproject.BoundText = "" Then
                    .rows = .FixedRows + RsDev.RecordCount
                Else
                    .rows = .FixedRows + RsDev.RecordCount - 1
                End If

                For i = .FixedRows To .rows - 1
                    .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("DEV_ID_Line_No").value), "", RsDev("DEV_ID_Line_No").value)
                    .TextMatrix(i, .ColIndex("CurrRow")) = IIf(IsNull(RsDev("CurrRow").value), 0, RsDev("CurrRow").value)
                    .TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(RsDev("Vatyo").value), 0, RsDev("Vatyo").value)
                    .TextMatrix(i, .ColIndex("FlgVat")) = IIf(IsNull(RsDev("FlgVat").value), "", RsDev("FlgVat").value)
                    .TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(RsDev("Vat").value), "", RsDev("Vat").value)
                    .TextMatrix(i, .ColIndex("UnitNo")) = IIf(IsNull(RsDev("UnitNo").value), "", RsDev("UnitNo").value)
            
                    .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), "", RsDev("DEV_ID_Line_No1").value)
            
                    .TextMatrix(i, .ColIndex("ExpensesID")) = get_Revenue_id(IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value))
            
                    .TextMatrix(i, .ColIndex("opr_fullcode")) = IIf(IsNull(RsDev("opr_fullcode").value), "", RsDev("opr_fullcode").value)
            
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
                    
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                    Else
                        .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                    End If
                    If val(RsDev!Vat & "") <> 0 Then
                        .TextMatrix(i, .ColIndex("DailyMonthly")) = val(RsDev!DailyMonthly & "")
                    Else
                         .TextMatrix(i, .ColIndex("DailyMonthly")) = ""
                    End If
                    .TextMatrix(i, .ColIndex("DailyMonthlyValue")) = IIf(IsNull(RsDev("DailyMonthlyValue").value), "", RsDev("DailyMonthlyValue").value)
                    'Double_Entry_Vouchers_Description
                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), "", RsDev("Double_Entry_Vouchers_Description").value)
            
                    '    .TextMatrix(I, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), _
                    '        "", RsDev("Account_Name").value)
                    
                       .TextMatrix(i, .ColIndex("ContItem")) = (IIf(IsNull(RsDev.Fields("ContItem").value), "", RsDev.Fields("ContItem").value))
                .TextMatrix(i, .ColIndex("PurchaseOrderNo")) = (IIf(IsNull(RsDev.Fields("PurchaseOrderNo").value), "", RsDev.Fields("PurchaseOrderNo").value))
                .TextMatrix(i, .ColIndex("LocationName")) = (IIf(IsNull(RsDev.Fields("LocationName").value), "", RsDev.Fields("LocationName").value))
                .TextMatrix(i, .ColIndex("RentType")) = (IIf(IsNull(RsDev.Fields("RentType").value), "", RsDev.Fields("RentType").value))
        
                    .TextMatrix(i, .ColIndex("Count")) = IIf(IsNull(RsDev("Count").value), "", RsDev("Count").value)
                    .TextMatrix(i, .ColIndex("price")) = IIf(IsNull(RsDev("price").value), "", RsDev("price").value)
                    .TextMatrix(i, .ColIndex("Discount")) = IIf(IsNull(RsDev("Discount").value), "", RsDev("Discount").value)
        
                    .TextMatrix(i, .ColIndex("MonthCount")) = RsDev!MonthCount & ""
                    .TextMatrix(i, .ColIndex("PurchOrderNo")) = RsDev!PurchOrderNo & ""
                    .TextMatrix(i, .ColIndex("CityFromId")) = RsDev!CityFromId & ""
                    .TextMatrix(i, .ColIndex("CityToId")) = RsDev!CityToId & ""
                    .TextMatrix(i, .ColIndex("FromCityName")) = RsDev!FromCityName & ""
                    .TextMatrix(i, .ColIndex("ToCityName")) = RsDev!ToCityName & ""
                    .TextMatrix(i, .ColIndex("des2")) = RsDev!des2 & ""
                        
        'Discount
                    .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
            
                    .TextMatrix(i, .ColIndex("Order_No")) = IIf(IsNull(RsDev("Order_No").value), "", RsDev("Order_No").value)
 
                    CarID = IIf(IsNull(RsDev("CarID").value), 0, RsDev("CarID").value)

                    If CarID <> 0 Then
                        GetCarName CarID, CarName
                        .TextMatrix(i, .ColIndex("CarId")) = IIf(IsNull(RsDev("CarID").value), "", RsDev("CarID").value)
             
                        .TextMatrix(i, .ColIndex("CarName")) = CarName
                 
                    End If

                    RsDev.MoveNext
                Next i

                      Me.XPTxtVal.text = Round(val(.Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))) - val(txtPayAmount), Decimal_Places)
                txtTotal = Round(val(Me.XPTxtVal.text), Decimal_Places)
                Me.XPTxtVal.text = Round(val(Me.XPTxtVal.text) - val(txtTotalFines), Decimal_Places)
                
                Me.XPTxtVal2.text = Round(val(Me.XPTxtVal.text) - .Aggregate(flexSTSum, .FixedRows, .ColIndex("Vat"), .rows - 1, .ColIndex("Vat")), Decimal_Places)
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

Private Sub SaveData(RowID, FileName)
    Dim Msg              As String
    Dim RsTemp           As New ADODB.Recordset
    Dim StrSQL           As String
    Dim BeginTrans       As Boolean
    Dim RsDev            As ADODB.Recordset
    Dim LngDevID         As Long
    Dim OtherInformation As New ClsGLOther
    Dim NoteID     As String
    'On Error GoTo ErrTrap
    Dim rsDummyAcc2      As New ADODB.Recordset
    Dim mAccFF           As String
    s = "Select A214 from branches "
    rsDummyAcc2.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummyAcc2.EOF Then
        mAccFF = Trim(rsDummyAcc2!A214 & "")
    End If
         
    If Me.TxtModFlg.text <> "R" Then
        If checkCustomerdata(val(Me.DCVendor.BoundText), val(XPTxtVal), val(DefaultInvoicetype.ListIndex), Dccurrency.text, Export) = False Then Exit Sub
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
                    Msg = "ÌÃ» ≈Œ Ì«— «·⁄„Ì·..!!"
                Else
                    Msg = "Select vendor..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCVendor.SetFocus
                Sendkeys "{F4}"
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
                Sendkeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 3 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
                Sendkeys "{F4}"
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
            '     If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
            '                 If SystemOptions.UserInterface = ArabicInterface Then
            '                     Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            '                 Else
            '                 Msg = "Cheque Due Date Not Valid...!!"
            '
            '                 End If
            '         MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '         DtpChequeDueDate.SetFocus
            '         SendKeys "{F4}"
            '         Exit Sub
            '     End If
        End If
    
        Dim i As Integer

        If CboPaymentType1.ListIndex = 0 Then

            With Fg_Journal

                For i = .FixedRows To .rows - 2

                    If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
                        '////////////////////////////////////////notes
               
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·«Ì ÌÊÃœ «Ì—«œ / Œœ„… ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                        Else
                            MsgBox "Select Expenses in line no" & i, vbCritical
                        End If

                        Exit Sub
              
                    End If
        
                Next i

            End With

            With Fg_Journal

                For i = .FixedRows To .rows - 2

                    If Not IsNumeric(.TextMatrix(i, .ColIndex("price"))) Or val(.TextMatrix(i, .ColIndex("price"))) < 0 Then
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
            
            Dim ISVAT As Boolean
            ISVAT = False
            With Fg_Journal
                For i = .FixedRows To .rows - 1
                    If val(.TextMatrix(i, .ColIndex("Vat"))) >= 0 Then
                        ISVAT = True
                    End If
                Next i
            End With
            Dim AccountVATDept As String
            If ISVAT = True And True = True Then
                If GetValueAddedAccount(XPDtbTrans.value, AccountVATDept) = False Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…"
                    Else
                        MsgBox "Value added account not specified"
                    End If
                    Exit Sub
                End If
            End If
            With Fg_Journal

                For i = .FixedRows To .rows - 2

                    If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Or val(.TextMatrix(i, .ColIndex("value"))) < 0 Then
                        '////////////////////////////////////////notes
               
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "«·’«›Ì ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ «ﬁ· „‰ ’›— " & i, vbCritical
                        Else
                            MsgBox "Enter Value in line no" & i, vbCritical
                        End If
               
                        Exit Sub
                    End If
        
                Next i

            End With

        End If
        'Õ”«»« 
        If Me.CboPaymentType1.ListIndex = 1 Then
      
            With Me.VSFlexGrid1

                For i = .FixedRows To .rows - 2

                    If .TextMatrix(i, .ColIndex("AccountCode")) = "" Then
                        '////////////////////////////////////////notes
               
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·«Ì ÌÊÃœ Õ”«» ›Ì «·”ÿ— —ﬁ„ " & i, vbCritical
                        Else
                            MsgBox "Select Expenses in line no" & i, vbCritical
                        End If

                        Exit Sub
              
                    End If
        
                Next i

            End With
   
            With Me.VSFlexGrid1

                For i = .FixedRows To .rows - 2

                    If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Or val(.TextMatrix(i, .ColIndex("value"))) <= 0 Then
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
        If Me.TxtModFlg.text = "N" Then
            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    '    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), XPDtbTrans.value) = False Then
                    '        Exit Sub
                    '    End If
                End If
            End If

        ElseIf Me.TxtModFlg.text = "E" Then

            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    '    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), XPDtbTrans.value, , , val(Me.XPTxtID.text)) = False Then
                    '        Exit Sub
                    '    End If
                End If
            End If
        End If
    
        Dim xrow As Integer

        With Fg_Journal

            For xrow = .rows - 1 To 2 Step -1

                If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

                    .rows = .rows - 1
                End If

            Next xrow

        End With
    
        With Me.VSFlexGrid1

            For xrow = .rows - 1 To 2 Step -1

                If .TextMatrix(xrow, .ColIndex("AccountCode")) = "" Then

                    .rows = .rows - 1
                End If

            Next xrow

        End With

        calcnets

        '-------------------------------------------------------------------------------------------
        '-------------------------------------------------------------------------------------------
        If TxtSerial.text = "" Then
            If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  "
                    Exit Sub
                Else
                    MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number "
                    Exit Sub
                End If

            Else
         
                If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  "
                        Exit Sub
                    Else
                        MsgBox "You must Define JE Coding "
                        Exit Sub
                    End If

                Else
                    TxtSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
                End If
            End If
        End If
 
        If TxtSerial1.text = "" Then
            If Voucher_coding(val(my_branch), XPDtbTrans.value, 8, 80) = "error" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   "
                    Exit Sub
                Else
                    MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number "
                    Exit Sub
                End If

            Else
                Dim ser As String
                ser = Voucher_coding(val(my_branch), XPDtbTrans.value, 63, 8063)
                If ser = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   "
                        Exit Sub
                    Else
                        MsgBox "  Enter Voucher No Manually or Define Coding "
                        Exit Sub
                    End If

                Else
                    TxtSerial1.text = ser 'Voucher_coding(val(my_branch), XPDtbTrans.value, 63, 8063)
                End If
            End If
        End If
        If mIsSaveFromExcel Then
            Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("des")) = Trim(Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("des"))) & ": " & TxtSerial1.text
            txt_general_des = Fg_Journal.TextMatrix(1, Fg_Journal.ColIndex("des"))
        End If
        Cn.BeginTrans
        BeginTrans = True
    
        '///////////////NOTESALL
        Dim A_NoteID As Long

        If TxtModFlg.text = "N" Then
            XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
            Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=85"))
            rs.AddNew
   
            Me.oldTxtSerial1.text = Trim$(Me.TxtSerial1.text)
 
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From notes Where notes_all=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
       
            If DcCostCenter.BoundText <> "" Then
                StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
        
        End If
    
    
    
        If AkarPayCheck.value = vbChecked Then
            rs("AkarPayCheck").value = 1

        Else
            rs("AkarPayCheck").value = 0
        End If
        '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("branch_no").value = val(Me.dcBranch.BoundText)
        rs("NoteID").value = val(XPTxtID.text)
        rs("bill_Type").value = Me.CboPaymentType1.ListIndex
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        rs("foxy_no").value = val(Text1.text)
        rs("order_no").value = TXT_order_no.text
        rs("CarID").value = val(Me.DcbCar.BoundText)
        
        rs("FromDate").value = txtFromDate.value
        rs("FromDateH").value = Me.txtFromDateH.value
       
        rs("ToDate").value = txtToDate.value
        rs("ToDateh").value = Me.txtToDateH.value
         rs("PayAmount").value = IIf(Me.txtPayAmount.text = "", Null, Trim(txtPayAmount.text))
          rs("Vendor").value = Trim(txtVendor.text)
        rs("OrderID").value = IIf(Me.txtOrderID.text = "", Null, Trim(txtOrderID.text))
        rs("Noteseril2").value = IIf(Me.TxtNoteSerial1.text = "", "", Trim(TxtNoteSerial1.text))
        rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, XPTxtVal.text)
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        rs("too").value = IIf(txtto.text = "", "", Trim(txtto.text))
        rs("general_des").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
        rs("TotalFines").value = IIf(txtTotalFines.text = "", Null, txtTotalFines.text)
        rs("RequestNo").value = IIf(txtRequestNo.text = "", Null, txtRequestNo.text)
        rs("ContractNo").value = IIf(txtContractNo.text = "", Null, txtContractNo.text)

        If ComResid(1).value = True Then
            rs.Fields("ComResid").value = 1
        Else
            rs.Fields("ComResid").value = 0
        End If
        rs("NewNO").value = IIf(TXTNewNO.text = "", Null, TXTNewNO.text)
        
        If CBoBasedON.ListIndex > -1 Then
            rs("BasedONID").value = CBoBasedON.ListIndex
        Else
            rs("BasedONID").value = 0
        End If
        
        

    rs("Currency_id").value = IIf(Dccurrency.BoundText = "", Null, val(Dccurrency.BoundText))
    rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.text), 1, txt_Currency_rate.text)
    rs("DateRec").value = txtDateRec.value
    rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
    rs("CIBAN").value = TXTIban.text
    rs("Invoicetype").value = Me.DefaultInvoicetype.ListIndex



        rs("txtManulaVat").value = val(txtManulaVat.text)
        rs("CusID").value = Null
        rs("NoteType").value = 85
        rs("NoteDate").value = XPDtbTrans.value
        rs("UserID").value = user_id
        rs("ExpensesID").value = IIf(XPCboExpensesType.text = "", Null, XPCboExpensesType.BoundText)
        If Me.CboPayMentType.ListIndex = 0 Then
            rs("BoxID").value = val(DcboBox.BoundText)
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("NoteCashingType").value = 0
        ElseIf Me.CboPayMentType.ListIndex = 1 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 1
    
        ElseIf Me.CboPayMentType.ListIndex >= 3 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = val(Me.CboPayMentType.ListIndex)
        
        ElseIf Me.CboPayMentType.ListIndex = 2 Then
            rs("NoteCashingType").value = 2
            rs("CusID").value = val(Me.DCVendor.BoundText)
        End If
         rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
        rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
        rs("Buy").value = "0"
        rs("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
        rs("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
        rs("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
 
        rs("OldNoteSerial1").value = Trim$(Me.oldTxtSerial1.text) '
     
        rs("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
        rs("numbering_type1").value = sand_numbering_type(63) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
     
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)

        If dcproject.BoundText <> "" Then
            '  rs("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
        Else
            '   rs("note_value_by_characters").value = WriteNo(Format(val(Me.XPTxtVal.text), "0.00"), 0, True, ".", , 0)
        End If

        If Me.TxtModFlg.text = "N" Then
            A_NoteID = CStr(new_id("Notes", "NoteID", "", True))
            TXT_A_NoteID.text = A_NoteID
        Else
            A_NoteID = val(TXT_A_NoteID.text)
        End If
    
        rs("A_NoteID").value = val(A_NoteID)
        If RowID <> "" Then
            rs("ExcelFile").value = FileName
            rs("ExcelRow").value = RowID
        End If
        
        
        
        rs.update
        '  Savetemp
        '/////////////////////Õ”«»«  ⁄«„Â
        Dim line_no As Integer

        If Me.CboPaymentType1.ListIndex = 1 Or VSFlexGrid1.rows > 1 Then
            If VSFlexGrid1.TextMatrix(1, VSFlexGrid1.ColIndex("AccountCode")) <> "" Then
            
            
            Set RsNotes = New ADODB.Recordset
            ' RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
            RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
            If TxtModFlg.text = "N" Then
           
            ElseIf Me.TxtModFlg.text = "E" Then
                '     StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.Text)
                '     Cn.Execute StrSQL, , adExecuteNoRecords
        
            End If
    
            '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
            ' rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
            ' rs("foxy_no").value = Val(Text1.text)
            'Õ”«»« 
            RsNotes.AddNew
             NoteID = CStr(new_id("Notes", "NoteID", "", True))
            RsNotes("NoteID").value = NoteID
            RsNotes("branch_no").value = val(Me.dcBranch.BoundText)
            RsNotes("order_no").value = TXT_order_no.text
            RsNotes("notes_all").value = Me.XPTxtID.text
            RsNotes("Note_Value").value = IIf(Not IsNumeric(XPTxtVal.text), 0, val(XPTxtVal.text))
            'RsNotes("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
            RsNotes("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
            RsNotes("too").value = IIf(txtto.text = "", "", Trim(txtto.text))
            '    RsNotes("general_des").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    
            If Me.CboPayMentType.ListIndex = 0 Then
                RsNotes("BoxID").value = val(DcboBox.BoundText)
                RsNotes("BankID").value = Null
                RsNotes("ChqueNum").value = Null
                RsNotes("DueDate").value = Null
                RsNotes("NoteCashingType").value = 0
            ElseIf Me.CboPayMentType.ListIndex = 1 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = 1
        
            ElseIf Me.CboPayMentType.ListIndex >= 3 Then
                RsNotes("BoxID").value = Null
                RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                RsNotes("NoteCashingType").value = val(Me.CboPayMentType.ListIndex)
           
            ElseIf Me.CboPayMentType.ListIndex = 2 Then
                RsNotes("CusID").value = val(DCVendor.BoundText)
            End If
    
            RsNotes("NoteType").value = 8063
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
            RsNotes("note_v_by_char_WithoutVat").value = Trim$(Me.lblValue2.Caption)
              
            RsNotes.update
            '„œÌ‰ Õ”«»« 
            With VSFlexGrid1
                line_no = 1
 
                For i = .FixedRows To .rows - 1
    
                    Dim project_id As Integer
    
                    If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                            
                                     Dim IntDEV_Type  As Integer
                        Dim SngDEV_Value As Single
                        line_no = line_no + 1
                        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                        OtherInformation.FlgVat = 0
                        OtherInformation.Vat = 0
                        OtherInformation.Vatyo = 0
                        OtherInformation.CurrRow = 0
                        OtherInformation.isExpens = True
                          
                        If ModAccounts.AddNewDev(LngDevID, line_no, DcboCreditSide.BoundText, val(.TextMatrix(i, .ColIndex("Value"))), 0, txt_general_des.text, NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , , val(Me.XPTxtID.text), , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                            GoTo ErrTrap
                        
                        End If
                                        
                        project_id = get_project_id(dcproject.BoundText, "REVENUE_account")
                        line_no = line_no + 1
                        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                        'OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                       ' OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                       ' OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                      '  OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                        If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("Value")), 1, .TextMatrix(i, .ColIndex("Des")), NoteID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , , , , , , val(.TextMatrix(i, .ColIndex("LineNo1"))), val(Me.XPTxtID.text), project_id, , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                            GoTo ErrTrap
                    
                        End If

                        line_no = line_no + 1
                        
                        OtherInformation.isExpens = False
                      
                    End If

                Next i

            End With

            'œ«∆‰ Õ”«»« 
      
        
            ' TxtModFlg.text = "R"
          '  GoTo ll
            End If
        End If
        
        Dim ExpensesID As Double
        
        Set RsNotes = New ADODB.Recordset
        '   RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
        RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
        Set RsDev = New ADODB.Recordset
        StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.* FROM         dbo.DOUBLE_ENTREY_VOUCHERS WHERE     (Double_Entry_Vouchers_ID = - 1)"
        RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

        '    «·ÿ—› «·„œÌ‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
        RsNotes.AddNew
        NoteID = CStr(new_id("Notes", "NoteID", "", True))
        RsNotes("NoteID").value = CStr(NoteID)
        RsNotes("branch_no").value = val(Me.dcBranch.BoundText)
 
        RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0)
        RsNotes("Remark").value = Me.txt_general_des
        RsNotes("foxy_no").value = val(Text1.text)

        If Me.CboPayMentType.ListIndex = 0 Then
            RsNotes("BoxID").value = val(DcboBox.BoundText)
            RsNotes("BankID").value = Null
            RsNotes("ChqueNum").value = Null
            RsNotes("DueDate").value = Null
            RsNotes("NoteCashingType").value = 0
        ElseIf Me.CboPayMentType.ListIndex = 1 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 1
        ElseIf Me.CboPayMentType.ListIndex >= 3 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = val(Me.CboPayMentType.ListIndex)
      
                            
        ElseIf Me.CboPayMentType.ListIndex = 2 Then
            RsNotes("CusID").value = val(DCVendor.BoundText)
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = Null
    
        End If
        ' RsNotes("order_no").value = txt_ORDER_NO.text
        '              RsNotes("CusID").value = Null
        RsNotes("NoteType").value = 8063
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
        RsNotes("note_v_by_char_WithoutVat").value = Trim$(Me.lblValue2.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
           
        RsNotes.update
        '«·ÿ—› «·„œÌ‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
        RsDev.AddNew
        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
        RsDev("branch_id").value = val(Me.dcBranch.BoundText)
        RsDev("DEV_ID_Line_No").value = line_no
        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
        RsDev("Account_Code").value = DcboCreditSide.BoundText
        RsDev("Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
        RsDev("Credit_Or_Debit").value = 0
        RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text  ' .TextMatrix(I, .ColIndex("des"))
        RsDev("RecordDate").value = Me.XPDtbTrans.value
        RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
                       
        RsDev("UserID").value = Me.DCboUserName.BoundText
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev("notes_all").value = Me.XPTxtID.text
        '   RsDev("project_id").value = project_id
                        
        RsDev.update
        'GoTo ll
        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
 
        line_no = line_no + 1

        '  «·«Ì—«œ«  œ«∆‰
    
        '//////////////////////////////////////Notes////////////////////////////////////

        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then

            '  RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            
            '«·ÿ—› «·„œÌ‰
Dim mAmount As Double
            With Fg_Journal

                line_no = 1
       
                project_id = get_project_id(dcproject.BoundText, "REVENUE_account")
                
                For i = .FixedRows To .rows - 1
   
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
                        RsNotes("branch_no").value = val(Me.dcBranch.BoundText)
                        RsNotes("Note_Value").value = .TextMatrix(i, .ColIndex("value"))
                        RsNotes("Count").value = val(.TextMatrix(i, .ColIndex("Count")))
                        RsNotes("price").value = val(.TextMatrix(i, .ColIndex("price")))
                        RsNotes("Discount").value = val(.TextMatrix(i, .ColIndex("Discount")))
                        
                        'Discount
                        RsNotes("ExpensesRemark").value = .TextMatrix(i, .ColIndex("des"))
                        
                        RsNotes("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))

                        RsNotes("foxy_no").value = val(Text1.text)
                        
                        RsNotes("DailyMonthly").value = val(.ValueMatrix(i, .ColIndex("DailyMonthly")))
                        RsNotes("DailyMonthlyValue").value = val(.ValueMatrix(i, .ColIndex("DailyMonthlyValue")))
                                   
                        If TXT_order_no.text <> "" Then
                            RsNotes("order_no").value = TXT_order_no.text
                        Else
                            RsNotes("order_no").value = IIf(.TextMatrix(i, .ColIndex("Order_No")) = "", Null, .TextMatrix(i, .ColIndex("Order_No")))
                        End If
            
                        RsNotes("CusID").value = Null
                        RsNotes("NoteType").value = 8063
                        RsNotes("NoteDate").value = XPDtbTrans.value
                        RsNotes("UserID").value = user_id
                        RsNotes("ExpensesID").value = .TextMatrix(i, .ColIndex("ExpensesID"))
                        
                        RsNotes("Count").value = val(.TextMatrix(i, .ColIndex("Count")))
                        RsNotes("price").value = val(.TextMatrix(i, .ColIndex("price")))
                        RsNotes("discount").value = val(.TextMatrix(i, .ColIndex("discount")))
                        
                        RsNotes("notes_all").value = Me.XPTxtID.text
                        RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
                        RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
                        RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
                        RsNotes("numbering_type1").value = sand_numbering_type(8) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
                
                        RsNotes("sanad_year").value = year(XPDtbTrans.value)
                        RsNotes("sanad_month").value = Month(XPDtbTrans.value)
                        RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(.TextMatrix(I, .ColIndex("value")), "0.00"), 0, True, ".")
                        RsNotes("note_v_by_char_WithoutVat").value = Trim$(Me.lblValue2.Caption) ' WriteNo(Format(.TextMatrix(I, .ColIndex("value")), "0.00"), 0, True, ".")
                        
                        RsNotes("MonthCount").value = val(.TextMatrix(i, .ColIndex("MonthCount")))
                        RsNotes("PurchOrderNo").value = val(.TextMatrix(i, .ColIndex("PurchOrderNo")))
                        RsNotes("CityFromId").value = val(.TextMatrix(i, .ColIndex("CityFromId")))
                        RsNotes("CityToId").value = val(.TextMatrix(i, .ColIndex("CityToId")))
                        RsNotes("des2").value = Trim(.TextMatrix(i, .ColIndex("des2")))
                
               
                        RsNotes("ContItem").value = (.TextMatrix(i, .ColIndex("ContItem")))
                        RsNotes("UnitNo").value = val(.TextMatrix(i, .ColIndex("UnitNo")))
                        
                        RsNotes("PurchaseOrderNo").value = (.TextMatrix(i, .ColIndex("PurchaseOrderNo")))
                        RsNotes("LocationName").value = (.TextMatrix(i, .ColIndex("LocationName")))
                        RsNotes("RentType").value = (.TextMatrix(i, .ColIndex("RentType")))
                         
                
                        RsNotes.update
              
                        '////////////////////////////////////////notes
 
                        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                        OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                        OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                        OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                        OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                        
                        If val(txtPayAmount) <> 0 Then
                            mAmount = val(.TextMatrix(i, .ColIndex("value"))) - ((val(.TextMatrix(i, .ColIndex("value"))) / (val(txtPayAmount) + val(Me.XPTxtVal.text)) * 100) / 100 * val(txtPayAmount))
                        Else
                            mAmount = val(.TextMatrix(i, .ColIndex("value")))
                        End If
                        'If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), val(.TextMatrix(i, .ColIndex("value"))), 1, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , val(.TextMatrix(i, .ColIndex("value"))), , , , , val(.TextMatrix(i, Fg_Journal.ColIndex("LineNo1"))), val(Me.XPTxtID.text), project_id, .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , , , , val(Me.dcBranch.BoundText), val(.TextMatrix(i, .ColIndex("CarId"))), , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), mAmount, 1, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , val(.TextMatrix(i, .ColIndex("value"))), , , , , val(.TextMatrix(i, Fg_Journal.ColIndex("LineNo1"))), val(Me.XPTxtID.text), project_id, .TextMatrix(i, Fg_Journal.ColIndex("opr_fullcode")), , , , , , val(Me.dcBranch.BoundText), val(.TextMatrix(i, .ColIndex("CarId"))), , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                            '   GoTo ErrTrap
                    
                        End If

                        line_no = line_no + 1
        
                    End If

                Next i

            End With
            If Me.dcproject.BoundText <> "" Then
           '     «·ÿ—› «·œ«∆‰   «Ì—«œ«  «·„‘—Ê⁄
                RsNotes.AddNew
                NoteID = CStr(new_id("Notes", "NoteID", "", True))
                RsNotes("NoteID").value = CStr(NoteID)
                RsNotes("branch_no").value = val(Me.dcBranch.BoundText)

                RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0)
                RsNotes("Remark").value = txt_general_des.text 'txtto.text

                If Me.CboPayMentType.ListIndex = 0 Then
                    RsNotes("BoxID").value = val(DcboBox.BoundText)
                    RsNotes("BankID").value = Null
                    RsNotes("ChqueNum").value = Null
                    RsNotes("DueDate").value = Null
                    RsNotes("NoteCashingType").value = 0
                ElseIf Me.CboPayMentType.ListIndex = 1 Then
                    RsNotes("BoxID").value = Null
                    RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                    RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                    RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                    RsNotes("NoteCashingType").value = 1

                ElseIf Me.CboPayMentType.ListIndex >= 3 Then
                    RsNotes("BoxID").value = Null
                    RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
                    RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
                    RsNotes("DueDate").value = Me.DtpChequeDueDate.value
                    RsNotes("NoteCashingType").value = val(Me.CboPayMentType.ListIndex)

                ElseIf Me.CboPayMentType.ListIndex = 2 Then
                    RsNotes("CusID").value = val(DCVendor.BoundText)
                    RsNotes("BoxID").value = Null
                    RsNotes("BankID").value = Null

                End If

                ' RsNotes("order_no").value = txt_ORDER_NO.text
                'RsNotes("CusID").value = Null
                RsNotes("NoteType").value = 8063
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
                RsNotes("note_v_by_char_WithoutVat").value = Trim$(Me.lblValue2.Caption)
                RsNotes.update
                
                project_id = get_project_id(dcproject.BoundText, "REVENUE_account")
                Set RsDev = New ADODB.Recordset
                
                 '  RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                
                line_no = line_no + 1

                With Fg_Journal

                    For i = .FixedRows To .rows - 1
        
                        If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                             '////////////////////////////////////////notes
                
                            If Not IsNumeric(.TextMatrix(i, .ColIndex("value"))) Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "·« Ì„ﬂ‰ « „«„ ⁄„·Ì… «·Õ›Ÿ ·⁄œ„ «œŒ«· ﬁÌ„… ›Ì «·”ÿ— —ﬁ„  " & i - 1, vbCritical: GoTo ErrTrap
                                 Else
                                    MsgBox "Cant save enter value in line :  " & i - 1, vbCritical: GoTo ErrTrap
                                 End If
               
                             End If

                             '  project_id = get_project_id(dcproject.BoundText, "REVENUE_account")
                            project_id = 0
                            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
                            OtherInformation.FlgVat = val(.TextMatrix(i, .ColIndex("FlgVat")))
                            OtherInformation.Vat = val(.TextMatrix(i, .ColIndex("Vat")))
                            OtherInformation.Vatyo = val(.TextMatrix(i, .ColIndex("Vatyo")))
                            OtherInformation.CurrRow = val(.TextMatrix(i, .ColIndex("CurrRow")))
                            If OtherInformation.FlgVat = 0 Then
                            If ModAccounts.AddNewDev(LngDevID, line_no, .TextMatrix(i, .ColIndex("AccountCode")), .TextMatrix(i, .ColIndex("value")) + val(.TextMatrix(i, .ColIndex("Vat"))), 0, .TextMatrix(i, .ColIndex("des")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , , .TextMatrix(i, .ColIndex("value")), , , , , setfoxy_Line, val(Me.XPTxtID.text), project_id, , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                GoTo ErrTrap
                    
                             End If

                            line_no = line_no + 1
                            End If
                         End If

                    Next i

                 End With
                 line_no = line_no + 1
                StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
                RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
                RsDev.AddNew
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
                RsDev("DEV_ID_Line_No").value = line_no
                RsDev("DEV_ID_Line_No1").value = setfoxy_Line
                RsDev("Account_Code").value = dcproject.BoundText
                RsDev("Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
                RsDev("Credit_Or_Debit").value = 1
                RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text ' .TextMatrix(I, .ColIndex("des"))
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)5
                       
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("notes_all").value = Me.XPTxtID.text
                 ' RsDev("project_id").value = project_id
                        
                RsDev.update
                 Dim sql As String
                sql = "Update notes    set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.text) & " and notetype=85" & "and NoteSerial1=" & TxtSerial1
                Cn.Execute sql
                sql = "Update   notes_all  set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.text) & " and notetype=85" & "and NoteSerial1=" & TxtSerial1
                Cn.Execute sql
                
                sql = "Update notes    set note_v_by_char_WithoutVat='" & WriteNo(Format(val(Me.XPTxtVal2.text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.text) & " and notetype=85" & "and NoteSerial1=" & TxtSerial1
                Cn.Execute sql
                sql = "Update   notes_all  set note_v_by_char_WithoutVat='" & WriteNo(Format(val(Me.XPTxtVal2.text) * 2, "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.text) & " and notetype=85" & "and NoteSerial1=" & TxtSerial1
                Cn.Execute sql

             End If
            If val(txtTotalFines) <> 0 Then
                line_no = line_no + 1
                 '„œÌ‰ €—«„« 
                RsDev.AddNew
                RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("DEV_ID_Line_No").value = line_no
                RsDev("DEV_ID_Line_No1").value = setfoxy_Line
                RsDev("Account_Code").value = mAccFF
                RsDev("Value").value = IIf(IsNumeric(txtTotalFines.text), txtTotalFines.text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = txt_general_des.text  ' .TextMatrix(I, .ColIndex("des"))
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
                       
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("notes_all").value = Me.XPTxtID.text
                 '   RsDev("project_id").value = project_id
                        
                RsDev.update
     
                 'GoTo ll
                 '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

             End If
             '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            LblDevID.Caption = LngDevID
            lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
         End If

ll:



 
            If SystemOptions.ApplyEinvoice Then savenewelectroncic
            
        
        If SystemOptions.IsBluee = True Then
 
   
                MsgBox SENDEINVOICE(Me.XPTxtID, True, val(Me.DCVendor.BoundText), , "notes_all", "NoteID"), vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  
        End If
        
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata
        
        If isCmdCreateTransa Then
             Exit Sub
         End If
    
        Select Case Me.TxtModFlg.text

             Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                 Else
                    Msg = " Saved... " & CHR(13)
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

                lbl(27).Caption = showLabel(TxtSerial1, oldTxtSerial1)
        
                Fg_Journal.Enabled = False
         End Select

         '«· Ê“Ì⁄ ⁄·Ï „—ﬂ“ «· ﬂ·›… «·⁄«„
   
         '     If Me.DcCostCenter.BoundText <> "" Then
        save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.text, "  ›« Ê—… Œœ„Ì…", Me.XPDtbTrans.value
         '     End If
        save_cost_center
         'Õ›Ÿ «·„’«—Ì› › ÃœÊ· «·„’«—Ì›
     
         '  If saveExpensesDetails(1, TxtSerial.text, TxtSerial1.text, txt_ORDER_NO.text, XPDtbTrans.value) = True Then
         '  End If
    
         'Õ›Ÿ »Ì«‰«  «·‘Ìﬂ« 
        saveChequeBoxContents1 (val(Me.XPTxtID.text))
    
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
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
         Else
            Msg = "cant save " & CHR(13)
            Msg = Msg + "Invalid entry value " & CHR(13)
            Msg = Msg + "Check data and try again"
    
         End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         Exit Sub
     End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
     Else
        Msg = "Sorr.... Error during saving " & CHR(13)
     End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Function save_cost_center()

    'on error resume next
    If Not IsNumeric(Text1.text) Then
        Exit Function
    End If
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql_str As String

    'Rs.Open "", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    sql_str = "select * from marakes_taklefa_temp where kedno=" & Text1.text
    rs.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs.RecordCount
        rs("ok").value = 1
        rs("NoteDate").value = XPDtbTrans.value
        rs("NoteSerial").value = TxtSerial.text
        rs("Remark").value = "   ›« Ê—… Œœ„Ì… —ﬁ„ " & TxtSerial1 & "    " & Me.txt_general_des
 
        rs.update
        rs.MoveNext
    Next i

End Function

Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND  kedno =" & val(Text1.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If

    'rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
StrSQL = "SELECT   *  from dbo.marakes_taklefa_temp Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
If CboPaymentType1.ListIndex = 0 Then
    With Fg_Journal
 
        .rows = .rows + 1

        For i = .FixedRows To .rows - 1
 
            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" And .TextMatrix(i, .ColIndex("FlgVat")) = "" Then
        
                rs.AddNew
                rs("cost_center_id").value = cost_center_id
                rs("cost_center").value = cost_center
                rs("value").value = val(.TextMatrix(i, .ColIndex("value")))
                rs("depit_or_credit").value = "œ«∆‰"
                rs("opr_id").value = Me.Text1.text
                rs("kedno").value = Me.Text1.text
                rs("opr_type").value = "›« Ê—… Œœ„Ì…"
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
          '      If val(.TextMatrix(i, .ColIndex("LineNo1"))) = 0 Then
                rs("line_no").value = val(.TextMatrix(i, .ColIndex("LineNo1")))
          '      Else
          '       rs("line_no").value = setfoxy_Line
          '      End If
                
                
                
                rs("record_date").value = record_date
                rs("general_des").value = 1
                rs.update
        
            End If

        Next i

    End With
Else

    With VSFlexGrid1
 
        .rows = .rows + 1

        For i = .FixedRows To .rows - 1

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
                rs("general_des").value = 1
                rs.update
        
            End If

        Next i

    End With



End If

    rs.Close
End Function

Private Sub Undo()
    Dim sgl As String
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            sgl = "delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
            Cn.Execute sgl, , adExecuteNoRecords
        
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)
         
        Case "E"
            sgl = "delete  marakes_taklefa_temp  where ok is null and  kedno =" & val(Text1.text)
            Cn.Execute sgl, , adExecuteNoRecords
        
            rs.Find "NoteID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

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
        If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
            Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–… «·⁄„·Ì…"
            Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    End If
    
    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
 
          '  StrSQL = "Delete From notes Where NoteID=" & val(TXT_A_NoteID.text)
             StrSQL = "Delete From notes Where notes_all=" & val(XPTxtID.text)
             
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From ExpensesDetails Where NoteSerial1='" & val(TxtSerial1.text) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
    
            StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
    
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
       
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Fg_Journal.Clear flexClearScrollable, flexClearEverything
                    Fg_Journal.rows = 3
                    Fg_Journal.Enabled = False
                
                    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                    VSFlexGrid1.rows = 2
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub

Function FillGridWithData()

End Function

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                .TextMatrix(i, .ColIndex("value")) = (val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("price")))) - val(.TextMatrix(i, .ColIndex("discount")))
                'Count  price
                '.TextMatrix(Row, .ColIndex("Count")) = 1
            End If

        Next i

    End With

    IntCounter = 0

    With Me.VSFlexGrid1

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

End Sub

Private Sub PutData()

    'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)
    With Fg_Journal

        If Len(TxtDes.text) > 0 Then
            .cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.text
            .cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
            .cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        Else
            .cell(flexcpData, .Row, .ColIndex("Des")) = ""
            .cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
            .cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        End If

    End With

End Sub

Function sand_numbering() As String
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    Dim auto_sanad_no As String
    Dim NO As String
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
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=85 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
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
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=85 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & mId(Format$(Now, "dd/mm/yyyy"), 4, 2)
            'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

            If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)

                If end_at = 0 Then end_at = NO + 1
                If NO >= end_at Then
                    sand_numbering = "error"
                    Exit Function
                End If
            End If

        Else

            If numbering_type = 3 Then
 
                detect_no.ConnectionString = connection_string
                detect_no.CommandType = adCmdText
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=85 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4)
                'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                detect_no.Refresh

                If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)

                    If end_at = 0 Then end_at = NO + 1
                    If NO >= end_at Then
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
                    auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & mId(Format$(Now, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & start_at

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
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (NO + 1)
                    '  End If
                      
                Else

                    If numbering_type = 3 Then
                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                        'no = 1
                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                        '    Else
                        NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                        auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (NO + 1)

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
    
    Dim X As Double
    X = CStr(new_id("foxy", "id1", "", True))
    setfoxy_Line = X
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id1").value = X ' last_line_id
 
    rs.update
    
End Function

Private Sub CBoBasedON_Change()

    'n
    With Me.Fg_Journal

        If Me.CBoBasedON.ListIndex = 0 Then

        ElseIf Me.CBoBasedON.ListIndex = 1 Then

            If SystemOptions.UserInterface = ArabicInterface Then
                lbl(21).Caption = "—ﬁ„ «·«„—"
            Else
                lbl(21).Caption = "  Order No"
            End If

        ElseIf Me.CBoBasedON.ListIndex = 2 Then

            If SystemOptions.UserInterface = ArabicInterface Then
                lbl(21).Caption = "—ﬁ„ «·›« Ê—… «·„»œ∆ÌÂ"
            Else
                lbl(21).Caption = "Performa Invoice NO"
            End If

        End If

        .TextMatrix(0, .ColIndex("order_no")) = lbl(21).Caption

    End With

End Sub

Function CuurentLogdata(Optional Currentmode As String)
'     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "—ﬁ„ «·›« Ê—… " & TxtSerial1.Text & CHR(13) & "   «· «—ÌŒ  " & XPDtbTrans & CHR(13) & "   «·›—⁄ " & Dcbranch & CHR(13) & "   „—ﬂ“ «· ﬂ·›… «·⁄«„  " & DcCostCenter & CHR(13) & "   ÿ—Ìﬁ… «·œ›⁄  " & CboPaymentType & CHR(13) & "   «·„‘—Ê⁄  " & dcproject & CHR(13) & "   «·„Ê—œ " & DCVendor & CHR(13) & "   «·Œ“Ì‰… " & DcboBox & CHR(13) & "   «·»‰ﬂ  " & DcboBankName & CHR(13) & "   —ﬁ„ «·‘Ìﬂ " & TxtChequeNumber & CHR(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & CHR(13) & "   —ﬁ„ ›« Ê—… «·„Ê—œ " & txtto & CHR(13) & "   »‰«¡ ⁄·Ï  " & CBoBasedON & "  »—ﬁ„  " & txt_ORDER_NO & CHR(13) & "   «·‘—Õ «·⁄«„  " & txt_general_des & CHR(13) & "   «Ã„«·Ì «·›« Ê—…    " & XPTxtValView
'        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Bill No " & TxtSerial1.Text & CHR(13) & "   Date  " & XPDtbTrans & CHR(13) & "   Branch " & Dcbranch & CHR(13) & "   CC  " & DcCostCenter & CHR(13) & "  Payment Type  " & CboPaymentType & CHR(13) & "   Project  " & dcproject & CHR(13) & "   Supplier " & DCVendor & CHR(13) & "   Box " & DcboBox & CHR(13) & "   Bank  " & DcboBankName & CHR(13) & "   Cheque No:   " & TxtChequeNumber & CHR(13) & "  Due Date  " & DtpChequeDueDate & CHR(13) & "  Supplier Bill No " & txtto & CHR(13) & "   Based On  " & CBoBasedON & "  No:  " & txt_ORDER_NO & CHR(13) & "  Remarks  " & txt_general_des & CHR(13) & "   Bill Total   " & XPTxtValView
'       If Currentmode <> "D" Then
'        AddToLogFile CInt(user_id), 8063, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TxtSerial, TxtSerial1
'    Else
'        AddToLogFile CInt(user_id), 8063, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtSerial, TxtSerial1
'    End If
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
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
        If Cmd(0).Enabled = False Then
            Exit Sub
        End If
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then
            Exit Sub
        End If
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then
            Exit Sub
        End If
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then
            Exit Sub
        End If
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then
            Exit Sub
        End If
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then
                Exit Sub
            End If
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
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    If BolRtl = True Then

        With TTP
            .Create Me.hWnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        End With

        With TTP
            .Create Me.hWnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·›Ê« Ì— «·„«·Ì…", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    Else

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "Add New Record..." & Wrap & "Shortcut Key F12 OR Enter" & Wrap & "OR Alt+N", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit the Current Record..." & Wrap & "Shortcut Key F11 " & Wrap & "OR Alt+E", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save the New Record OR Save the Editing in the Current Record..." & Wrap & "Shortcut Key F10 " & Wrap & "OR Alt+S", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Cancel the New Record OR Cancel Editing in the Current Record..." & Wrap & "Shortcut Key F9 " & Wrap & "OR Alt+U", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete the Current Record..." & Wrap & "Shortcut Key F8 " & Wrap & "OR Alt+D", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Close this Screen" & Wrap & "OR Alt+X", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "Display Help for this Screen" & Wrap & "Shortcut Key F1" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
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
                SaveData "", ""

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPCboExpensesType_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("ExpensesType", "ID", val(Me.XPCboExpensesType.BoundText))
    End If

End Sub

Private Sub XPDtbTrans_Change()

    If Me.TxtModFlg = "E" Then
        If Month(rs("NoteDate").value) = Month(XPDtbTrans.value) Then
            Exit Sub
        End If
    End If

    If Trim(TxtSerial1.text) <> "" Then
        oldTxtSerial1.text = TxtSerial1.text
    End If

    TxtSerial.text = ""
    TxtSerial1.text = ""

End Sub

Private Sub XPTxtVal_Change()
    'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0)
    XPTxtValView.text = Format(val(XPTxtVal.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
    XPTxtValView2.text = Format(val(XPTxtVal2.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 0)
        Me.lblValue2.Caption = WriteNo(Format(Me.XPTxtVal2.text, "0.00"), 0, True, ".", , 0)

    Else

        'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0, , , , 1)
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 1)
        Me.lblValue2.Caption = WriteNo(Format(Me.XPTxtVal2.text, "0.00"), 0, True, ".", , 1)

    End If
    
End Sub

Private Sub XPTxtVal2_Change()
    'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0)
    
    XPTxtValView2.text = Format(val(XPTxtVal2.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

    If SystemOptions.UserInterface = ArabicInterface Then
   
        Me.lblValue2.Caption = WriteNo(Format(Me.XPTxtVal2.text, "0.00"), 0, True, ".", , 0)

    Else

        'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0, , , , 1)
  
        Me.lblValue2.Caption = WriteNo(Format(Me.XPTxtVal2.text, "0.00"), 0, True, ".", , 1)

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
    Dim FG As VSFlex8UCtl.VSFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    'Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set FG = FrmView.vsfGroup1.VSFlexGrid

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
    
        StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & "Remark, UserName From ExpensesReport"
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
    FrmView.vsfGroup1.VSFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "Note_Value"
    FrmView.vsfGroup1.sql = StrSQL
    FrmView.vsfGroup1.ShowTreeGroups = True
    FrmView.vsfGroup1.update
    FrmView.SetDblClickRetrun Me, "NoteID"
    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·»Ì«‰«  «·„’—Ê›« "
    FrmView.show
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
    lbl(24).Caption = "Hint"
    CmdAttach.Caption = "Attachments"
    lbl(28).Caption = "Car"
    lbl(25).Caption = "This Window Allow To Refister Servic Invoice"
    lbl(23).Caption = "Invoice Type"
    Label3.Caption = "GL No."
    lbl(14).Caption = "Project#"
    'Label1.Caption = "Manual #"
    Me.ALLButton1.Caption = "Cost Center"
    lbl(15).Caption = "Payment Method"
    lbl(16).Caption = "Box Name"
    lbl(20).Caption = "General Des"
    lbl(21).Caption = "Order No:"

    Label8.Caption = "General C. C."
    Label1.Caption = "Branch"
    lbl(26).Caption = "Based ON"

    With Me.CBoBasedON
        .Clear
        .AddItem "Without"
        .AddItem "Purchase Invoices"
        .AddItem "Performa Invoices"
        .AddItem "Production Order"
    End With

    With Me.CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
        .AddItem "Credit"
        .AddItem "P Cheque "
    End With

    With Me.CboPaymentType1
        .Clear
        .AddItem "Expenses"
        .AddItem "Accounts"
     
    End With

    CmdRemove.Caption = "Delete Row"
    Me.Caption = "Service Invoice"
    Me.Ele.Caption = Me.Caption

    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    Me.lbl(4).Caption = " Vchr#"
    Me.lbl(1).Caption = " Date"
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
Label10.Caption = "Manual No."

    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("CarName")) = "Car Name"
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = " Expenses Name"
        .TextMatrix(0, .ColIndex("value")) = "Net"
        .TextMatrix(0, .ColIndex("des")) = "description"
        .TextMatrix(0, .ColIndex("opr_fullcode")) = "Operation"
        .TextMatrix(0, .ColIndex("count")) = "Qty"
        .TextMatrix(0, .ColIndex("price")) = "Price"
        .TextMatrix(0, .ColIndex("Discount")) = "Discount"
        .TextMatrix(0, .ColIndex("Vatyo")) = "VAT %"
        .TextMatrix(0, .ColIndex("Vat")) = "VAT"
    End With

    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = " Account Name"
        .TextMatrix(0, .ColIndex("Account_Serial")) = " Account Code  "
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("Des")) = "Description"
    End With

End Sub

Private Sub dcVendor_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 182
        FrmCustemerSearch.show vbModal

    End If
 

    If KeyCode = vbKeyF5 Then
        Dcombos.GetCustomersSuppliers 1, Me.DCVendor
    End If
    
End Sub




Sub Savetemp()
    
    
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
 
 
SaveQRCode6 "notes_all", "NoteID", val(XPTxtID.text), TxtSerial1.text, (XPDtbTrans.value), _
        (XPTxtValView.text), Picture1, 0, (txtVatTotal.text), (XPTxtValView.text), val(dcBranch.BoundText)


End Sub











Function savenewelectroncic()
   'vat data
    Dim InvoiceTypeCodeID As Integer
    rs("CIBAN").value = TXTIban.text
    'vat data
      rs("RecTime").value = Time
            
   
   
  If val(DCDocTypes.BoundText) <> 0 Then
  'wAEL
    getDocAccounts val(DCDocTypes.BoundText), , , , , , , , , , , , InvoiceTypeCodeID
  Else
 InvoiceTypeCodeID = 388
  End If
  InvoiceTypeCodeID = 388
  rs("InvoiceTypeCodeID").value = InvoiceTypeCodeID
 
 
 
 If val(Me.DefaultInvoicetype.ListIndex) = 0 Then
   
   
    If Export = 1 Then
    rs("InvoiceTypeCodename").value = "0100100"
    Else
      rs("InvoiceTypeCodename").value = "0100000"
   End If
   
   
   
   
   Else
    rs("InvoiceTypeCodename").value = "0200000"
   End If

   rs("DocumentCurrencyCode").value = Dccurrency.text
   rs("TaxCurrencyCode").value = Dccurrency.text
  rs("ActualDeliveryDate").value = txtDateRec.value
 rs("LatestDeliveryDate").value = txtDateRec.value
Dim PaymentMeansCode As String
         
            '10 In cash
            '30 Credit
            '42 Payment to bank account
            '48 Bank card
            '1 Instrument not defined(Free text)
            Dim paymentnote
        If CboPayMentType.ListIndex = 0 Then ' ‰ﬁœ«
                  PaymentMeansCode = "10"
                      paymentnote = "Payment by Cash"
        ElseIf CboPayMentType.ListIndex = 1 Then ' ¬Ã·
                 PaymentMeansCode = "30"
                 paymentnote = "Payment by Credit"
         ElseIf CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then  '  ÕÊÌ· »‰ﬂÌ
                    If SystemOptions.AllowSalesMultyPayed = True Then
                     PaymentMeansCode = "48" 'ﬂ«— 
                      paymentnote = "Payment by Bank Card"
                    Else
                    PaymentMeansCode = "42" '»‰ﬂ Õ”«»
                    paymentnote = "Payment by Bank Account"
                    End If
         
         End If
         
         rs("PaymentMeansCode").value = PaymentMeansCode
      
rs("paymentnote").value = paymentnote
rs.update
End Function

Private Sub DcCurrency_Change()

    If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    If Me.Dccurrency.BoundText <> "" Then
        txt_Currency_rate.text = get_currency_rate(Me.Dccurrency.BoundText)
    Else
        txt_Currency_rate.text = 1
    End If

End Sub

Private Sub DcCurrency_Click(Area As Integer)
    DcCurrency_Change
End Sub





Private Sub Delay(Seconds As Integer)
    Sleep Seconds * 150
End Sub


