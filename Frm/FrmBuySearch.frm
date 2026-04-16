VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmBuySearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ⁄„·Ì… ‘—«¡"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10650
   Icon            =   "FrmBuySearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtPlatNo 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   750
      TabIndex        =   82
      Top             =   3420
      Width           =   1980
   End
   Begin VB.TextBox txtPrevValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Height          =   405
      Left            =   1080
      TabIndex        =   80
      Top             =   4470
      Width           =   1200
   End
   Begin VB.TextBox txtContainerNo 
      BackColor       =   &H0000FFFF&
      Height          =   345
      Left            =   4380
      TabIndex        =   78
      Top             =   3060
      Width           =   1155
   End
   Begin VB.TextBox order_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4560
      TabIndex        =   76
      Top             =   3360
      Width           =   1830
   End
   Begin VB.TextBox txtOldNoteSerial1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   180
      TabIndex        =   74
      Top             =   4140
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.TextBox txtPurchOrderNo 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   180
      TabIndex        =   72
      Top             =   3780
      Width           =   1320
   End
   Begin VB.TextBox TxtTransactionComment 
      Alignment       =   1  'Right Justify
      Height          =   615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   6390
      Width           =   3015
   End
   Begin VB.TextBox Txtorder_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   2400
      Width           =   2625
   End
   Begin VB.TextBox TxtCashCustomerPhone 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   3120
      Width           =   2625
   End
   Begin VB.TextBox TxtCashCustomerName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   2760
      Width           =   2625
   End
   Begin VB.TextBox TxtNetValue 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8070
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   4680
      Width           =   1425
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   4680
      Width           =   3495
      Begin XtremeSuiteControls.RadioButton RdNet 
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   53
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   ">"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdNet 
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   54
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "<"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdNet 
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   55
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "="
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdNet 
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   56
         Top             =   0
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   ">="
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdNet 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   57
         Top             =   0
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "<="
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   4200
      Width           =   3495
      Begin XtremeSuiteControls.RadioButton RdTotal 
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   47
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   ">"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdTotal 
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   48
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "<"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdTotal 
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   49
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "="
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdTotal 
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   50
         Top             =   0
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   ">="
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdTotal 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   51
         Top             =   0
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "<="
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.TextBox TxtTotalValue 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8070
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   4200
      Width           =   1425
   End
   Begin VB.TextBox txtManualNO 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4590
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   2400
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "‰ﬁ· «·»ÕÀ «·Ì «·”‰œ"
      Height          =   375
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   5160
      Width           =   1575
   End
   Begin ImpulseButton.ISButton CmdShowMoreOptions 
      Height          =   255
      Left            =   7500
      TabIndex        =   7
      Top             =   5160
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   450
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
      ButtonImage     =   "FrmBuySearch.frx":030A
      ColorButton     =   14871017
      ColorHoverText  =   12582912
      ButtonToggles   =   1
      DrawFocusRectangle=   0   'False
      RightToLeft     =   -1  'True
      ButtonImageToggled=   "FrmBuySearch.frx":06A4
      ColorToggledHoverText=   192
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·›« Ê—…  Õ ÊÏ ⁄·Ï Â–« «·’‰›"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1395
      Index           =   1
      Left            =   4170
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   5580
      Width           =   6495
      Begin VB.TextBox TxtItemCode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   1275
      End
      Begin VB.CheckBox ChkSerialSearchType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "»ÕÀ „ÿ«»ﬁ"
         Height          =   285
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   990
         Width           =   1455
      End
      Begin VB.TextBox TxtItemSerial 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   3315
      End
      Begin VB.TextBox TxtItemPrice 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox TxtItemQty 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo DCboItem 
         Height          =   315
         Left            =   690
         TabIndex        =   9
         Top             =   240
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton CmdItemSearch 
         Height          =   345
         Left            =   210
         TabIndex        =   35
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBuySearch.frx":0A3E
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ﬂÊœ «·’‰›"
         Height          =   345
         Index           =   6
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   690
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "”⁄— «·’‰›"
         Height          =   315
         Index           =   5
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   615
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "”Ì—Ì«· "
         Height          =   315
         Index           =   4
         Left            =   5610
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1020
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ﬂ„Ì… «·’‰›"
         Height          =   315
         Index           =   3
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   645
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "«”„ «·’‰›"
         Height          =   315
         Index           =   2
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.ComboBox CboPaymentType 
      Height          =   315
      Left            =   180
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4860
      Width           =   2625
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì «·› —…"
      ForeColor       =   &H00FF0000&
      Height          =   705
      Index           =   0
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   5190
      Width           =   4005
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   345
         Left            =   1980
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   202309633
         CurrentDate     =   38979.743287037
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   375
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   202309633
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   255
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   10
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   315
         Width           =   345
      End
   End
   Begin VB.TextBox TxtVal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   6600
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox XPTxtClientsName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Height          =   315
      Left            =   60
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1980
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox XPTxtBillNum 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7695
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2400
      Width           =   1785
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2325
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10605
      _cx             =   18706
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   23
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmBuySearch.frx":0FD8
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
   Begin VB.CheckBox XPChkSearchType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì· »«·ﬂ«„· ›ﬁÿ"
      Height          =   345
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2370
      Visible         =   0   'False
      Width           =   795
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   2100
      TabIndex        =   15
      Top             =   5970
      Width           =   1005
      _ExtentX        =   1773
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
      Left            =   1125
      TabIndex        =   16
      Top             =   5970
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   330
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5970
      Width           =   735
      _ExtentX        =   1296
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
   Begin MSDataListLib.DataCombo DcboStores 
      Height          =   315
      Left            =   6150
      TabIndex        =   2
      Top             =   3120
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboUsers 
      Height          =   315
      Left            =   7230
      TabIndex        =   3
      Top             =   3450
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdCusSearch 
      Height          =   345
      Index           =   0
      Left            =   3840
      TabIndex        =   40
      Top             =   2760
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmBuySearch.frx":1382
      DrawFocusRectangle=   0   'False
   End
   Begin MSDataListLib.DataCombo DCboClientsName 
      Height          =   315
      Left            =   4590
      TabIndex        =   42
      Top             =   2760
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCEquipments 
      Height          =   315
      Left            =   5670
      TabIndex        =   66
      Top             =   3840
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "6"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCPaymentNet 
      Height          =   315
      Left            =   2430
      TabIndex        =   68
      Top             =   3780
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "6"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label LblPla 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «··ÊÕ…"
      Height          =   210
      Left            =   2520
      TabIndex        =   83
      Top             =   3450
      Width           =   1200
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁ—«¡Â «·Õ«·ÌÂ"
      Height          =   195
      Index           =   163
      Left            =   2070
      TabIndex        =   81
      Top             =   4545
      Width           =   1200
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«—«„ﬂÊ"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   5040
      TabIndex        =   79
      Top             =   3120
      Width           =   1590
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»‰«¡ ⁄·Ì"
      Height          =   285
      Index           =   12
      Left            =   6540
      TabIndex        =   77
      Top             =   3390
      Width           =   660
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·›« Ê—… «·«’·Ï"
      Height          =   285
      Index           =   9
      Left            =   2280
      TabIndex        =   75
      Top             =   4170
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«„— «·‘—«¡"
      Height          =   255
      Index           =   137
      Left            =   1530
      TabIndex        =   73
      Top             =   3810
      Width           =   750
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "„·«ÕŸ« "
      Height          =   315
      Index           =   8
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   6510
      Width           =   705
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«‰Ê«⁄ «·œ›⁄"
      Height          =   285
      Index           =   12
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   3810
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„⁄œÂ/«·”Ì«—…"
      Height          =   240
      Index           =   62
      Left            =   9570
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   3840
      Width           =   1050
   End
   Begin VB.Label XPLbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «„— «·»Ì⁄"
      Height          =   315
      Index           =   11
      Left            =   2775
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   2400
      Width           =   1305
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ·Ì›Ê‰ "
      Height          =   315
      Index           =   10
      Left            =   2415
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   3120
      Width           =   1305
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄„Ì· ‰ﬁœÌ"
      Height          =   315
      Index           =   9
      Left            =   2415
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   2760
      Width           =   1305
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌ„… «·’«›Ì…"
      Height          =   315
      Index           =   8
      Left            =   9570
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   4680
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌ„… «·«Ã„«·Ì"
      Height          =   315
      Index           =   7
      Left            =   9585
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   4170
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
      Height          =   315
      Index           =   6
      Left            =   6285
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   2370
      Width           =   1305
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   315
      Index           =   7
      Left            =   9570
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰ ÌÃ… «·»ÕÀ:"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   4740
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   5130
      Width           =   2685
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
      Height          =   285
      Index           =   5
      Left            =   2820
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   4920
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Œ“‰"
      Height          =   315
      Index           =   0
      Left            =   9570
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   3090
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄—÷"
      Height          =   315
      Index           =   4
      Left            =   11760
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”‰œ"
      Height          =   315
      Index           =   3
      Left            =   9570
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2370
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… «·⁄—÷"
      Height          =   315
      Index           =   2
      Left            =   4710
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   5160
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·›« Ê—…"
      Height          =   315
      Index           =   1
      Left            =   11280
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì·"
      Height          =   315
      Index           =   0
      Left            =   9570
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2730
      Width           =   1065
   End
End
Attribute VB_Name = "FrmBuySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim cSearchDcbo(3) As clsDCboSearch
Public invoiceSerach As Boolean
Public TypeInvoice As Long

Public Index As Integer
Private m_DealingForm As GridTransType
Dim M_ExtraRetrunObject As Object
Dim M_ExtraRetrunObject1 As Object
Dim M_ExtraRetrunObject2 As Object
Public localindex As Integer
Public RetrunFrm As Form
Dim allTransaction_ID As String
Public mCusId As Integer
Public mmItemId As Long

Private Sub CboPayMentType_Change()
    DCPaymentNet.Visible = False
    XPLbl(12).Visible = False
If DealingForm = InvoiceTransaction And val(CboPayMentType.ListIndex) = 3 Then
    DCPaymentNet.Visible = True
    XPLbl(12).Visible = True
  End If
End Sub

Private Sub CboPayMentType_Click()
CboPayMentType_Change
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim Msg As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.rows = 2
           
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                Else
                    Msg = "No Results Found"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(1).Caption = "‰ ÌÃ… «·»ÕÀ: " & rs.RecordCount
            Else
                Me.lbl(1).Caption = "  Search Result: " & rs.RecordCount
            End If
       
            Retrive

        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 2

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = Msg + "Search critiria error" & CHR(13)
         
        End If
        
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub CmdCusSearch_Click(Index As Integer)
Select Case Me.DealingForm
 

        Case salespricelist
   
        
              Load FrmCustemerSearch
            FrmCustemerSearch.SearchType = 1
            FrmCustemerSearch.RetrunType = 1
            Set FrmCustemerSearch.DcboCustomers = Me.DCboClientsName
            FrmCustemerSearch.show vbModal
      
      
           
        Case ShowPrice
 
        
              Load FrmCustemerSearch
            FrmCustemerSearch.SearchType = 1
            FrmCustemerSearch.RetrunType = 1
            Set FrmCustemerSearch.DcboCustomers = Me.DCboClientsName
            FrmCustemerSearch.show vbModal
            
        Case InvoiceTransaction
    
              Load FrmCustemerSearch
            FrmCustemerSearch.SearchType = 1
            FrmCustemerSearch.RetrunType = 1
            Set FrmCustemerSearch.DcboCustomers = Me.DCboClientsName
            FrmCustemerSearch.show vbModal
   
        Case ReturnSalling
        
              Load FrmCustemerSearch
            FrmCustemerSearch.SearchType = 1
            FrmCustemerSearch.RetrunType = 1
            Set FrmCustemerSearch.DcboCustomers = Me.DCboClientsName
            FrmCustemerSearch.show vbModal
            
            
        Case PurchaseTransaction
          
             Load FrmCustemerSearch
            FrmCustemerSearch.SearchType = 2
            FrmCustemerSearch.RetrunType = 1
            Set FrmCustemerSearch.DcboCustomers = Me.DCboClientsName
            FrmCustemerSearch.show vbModal
            
            
           Case InventoryOut
          
             Load FrmCustemerSearch
            FrmCustemerSearch.SearchType = 2
            FrmCustemerSearch.RetrunType = 1
            Set FrmCustemerSearch.DcboCustomers = Me.DCboClientsName
            FrmCustemerSearch.show vbModal
               
End Select
End Sub

Private Sub CmdItemSearch_Click()
    Load FrmItemSearch
    FrmItemSearch.RetrunType = 1
    Set FrmItemSearch.DcboItems = Me.DCboItem
    PutFormOnTop Me.hWnd, False
    FrmItemSearch.show vbModal
    PutFormOnTop Me.hWnd, True
End Sub

Private Sub CmdShowMoreOptions_Click()
CmdShowMoreOptions.value = True
    If CmdShowMoreOptions.value = True Then
        Me.Fra(1).Visible = True
        'Me.Height = Me.Fra(1).top + Fra(1).Height + 400
        Me.Height = Me.Fra(1).top + Fra(1).Height + 500 ' GetMyTitleBarHight(Me.hwnd)
        'Me.Height = Me.ScaleHeight
    Else
        Me.Fra(1).Visible = False
        Me.Height = Me.Fra(1).top + 500
    End If

End Sub

Private Sub Command1_Click()
  

            
Dim StrSQL As String
If allTransaction_ID = "" Then
    Exit Sub
End If
allTransaction_ID = mId(allTransaction_ID, 2, Len(allTransaction_ID))
StrSQL = "select * from Transactions where Transaction_ID in (" & allTransaction_ID & ")"
  Select Case Me.DealingForm
            Case InvoiceTransaction
            If localindex = 0 Then
            frmsalebill.generalSearch (StrSQL)
            frmsalebill.invoiceSerach = True
            ElseIf localindex = 4 Then
            frmsalebill4.generalSearch (StrSQL)
            frmsalebill4.invoiceSerach = True
            ElseIf localindex = 3 Then
            frmsalebill3.generalSearch (StrSQL)
            frmsalebill3.invoiceSerach = True
            ElseIf localindex = 4 Then
            frmsalebill6.generalSearch (StrSQL)
            frmsalebill6.invoiceSerach = True
            ElseIf localindex = 5 Then
            frmsalebill6.generalSearch (StrSQL)
            frmsalebill6.invoiceSerach = True
            
            End If
            
        
        
        Case InvoiceTransactionCompose
  '      frmsalebillCompose.generalSearc (StrSQL)
            
        frmsalebillCompose.invoiceSerach = True
Case InventoryOut
FrmOut.generalSearch (StrSQL)

 End Select
 
End Sub

Private Sub fg_Click()
    Dim StrSQL As String
    Dim Num As Integer
    Dim RowNum As Integer
    Dim StrQry As String
    Dim RsDetails As ADODB.Recordset
    Dim DateTemp As Date
    Dim Msg As String

    On Error GoTo ErrTrap
 
    If Not FG.TextMatrix(FG.Row, 1) = "" Then
If Index <> 1 Then

If Index = 2 Then
 FrmMoving.TXT_order_no.text = FG.TextMatrix(FG.Row, 3)
ElseIf Index = 3 Then

  FrmPO4.TxtOrder.text = FG.TextMatrix(FG.Row, 3)
ElseIf Index = 310 Then
        
        FrmCarAuthontication.txtSalesInvoiceOrder.text = FG.TextMatrix(FG.Row, 3)
        FrmCarAuthontication.TxtCusID.text = val(FG.TextMatrix(FG.Row, FG.ColIndex("CusId")))
        FrmCarAuthontication.cmbItems.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("ItemId")))
        'FrmCarAuthontication.TxtClientCode.Text = Trim(fg.TextMatrix(fg.Row, fg.ColIndex("Fullcode")))
        
        FrmCarAuthontication.TxtCliientName.text = Trim(FG.TextMatrix(FG.Row, FG.ColIndex("ClientNmae")))
        FrmCarAuthontication.retInfoCustomer
    '    FrmCarAuthontication.txtAddres.Text = Trim(fg.TextMatrix(fg.Row, fg.ColIndex("Address")))
        
        If Trim(FG.TextMatrix(FG.Row, FG.ColIndex("Cus_mobile"))) = "" Then
    '            FrmCarAuthontication.TxtClientPhone.Text = Trim(fg.TextMatrix(fg.Row, fg.ColIndex("CashCustomerPhone")))
        Else
    '        FrmCarAuthontication.TxtClientPhone.Text = Trim(fg.TextMatrix(fg.Row, fg.ColIndex("Cus_mobile")))
        End If
       

ElseIf Index = 4 Then
 FrmInpout.TXT_order_no.text = FG.TextMatrix(FG.Row, 3)

ElseIf Index = 222 Then
FrmProductionOrder.TXT_order_no.text = FG.TextMatrix(FG.Row, 3)

ElseIf Index = 223 Then
FrmProductionOrder.TXT_order_no.text = FG.TextMatrix(FG.Row, 3)

ElseIf Index = 5 Then
FrmReturnpurchases.TxtInvID.text = val(FG.TextMatrix(FG.Row, 1))
FrmReturnpurchases.TxtInvSerial.text = FG.TextMatrix(FG.Row, 3)
ElseIf Index = 6 Then

 FrmPO8.TxtPO6.text = FG.TextMatrix(FG.Row, 3)

ElseIf Index = 62 Then
   FrmPO11.TxtPO6.text = FG.TextMatrix(FG.Row, 3)
ElseIf Index = 7 Then

FrmOut.TXT_order_no.text = FG.TextMatrix(FG.Row, 3)
ElseIf Index = 8 Or Index = 9 Then
  FrmProductionPlan.TxtNoteSerial.text = FG.TextMatrix(FG.Row, 3)


ElseIf Index = 10 Then
   FrmOut1.TXT_order_no.text = FG.TextMatrix(FG.Row, 3)
ElseIf Index = 11 Then
 FrmDestruction.TXT_order_no.text = val(FG.TextMatrix(FG.Row, 1))
 FrmDestruction.TxtNoteSerial1.text = FG.TextMatrix(FG.Row, 3)

ElseIf Index = 12 Then
FrmOut.TXT_order_no.text = FG.TextMatrix(FG.Row, 3)
ElseIf Index = 13 Then
 FrmReturnSalling.TxtInvSerial.text = FG.TextMatrix(FG.Row, 3)
ElseIf Index = 658 Then
 FrmInpout.TXT_order_no.text = FG.TextMatrix(FG.Row, 3)
 ElseIf Index = 131 Then
 frmsalebill2.TxtInvSerial.text = FG.TextMatrix(FG.Row, 3)
' frmsalebill2.RetriveReSalin
 ElseIf Index = 1205 Then
 frmsalebill2.TXTPrintInvoice.text = FG.TextMatrix(FG.Row, 3)
  frmsalebill2.printInvoice
  
ElseIf Index = 14 Then
    FrmShipmentOrder.TxtPONo.text = FG.TextMatrix(FG.Row, 3)
ElseIf Index = 15 Then

 FrmPO10.TxtPO6.text = FG.TextMatrix(FG.Row, 3)


ElseIf Index = 16 Then

 FrmPO10.TxtPO6.text = FG.TextMatrix(FG.Row, 3)

ElseIf Index = 17 Then

   FrmTypeExchange.TxtOrderNo.text = FG.TextMatrix(FG.Row, 3)
    FrmTypeExchange.txtTransaction_ID.text = FG.TextMatrix(FG.Row, 1)
       FrmTypeExchange.TxtPrice = gettransactiontotal(val(FrmTypeExchange.txtTransaction_ID.text))


 ElseIf Index = 18 Or Index = 19 Then

    FrmTypeExchange.TxtOrderNo.text = FG.TextMatrix(FG.Row, 3)
    FrmTypeExchange.txtTransaction_ID.text = FG.TextMatrix(FG.Row, 1)
    FrmTypeExchange.DCboCashType121.ListIndex = 1
    FrmTypeExchange.DBCboClientName.text = FG.TextMatrix(FG.Row, FG.ColIndex("ClientNmae"))
    FrmTypeExchange.TxtPrice = gettransactiontotal(val(FrmTypeExchange.txtTransaction_ID.text))
 ElseIf Index = 20 Then
 FrmPO3.TxtPONo.text = FG.TextMatrix(FG.Row, 3)

 ElseIf Index = 21 Then
 FrmInpout.TXT_order_no.text = FG.TextMatrix(FG.Row, 3)
   
Else
        Select Case Me.DealingForm
  
            Case PurchaseTransaction

                If Me.ExtraRetrunObject Is Nothing Then
                    RetrunFrm.Retrive val(FG.TextMatrix(FG.Row, 1))
                Else
                    Me.ExtraRetrunObject = (FG.TextMatrix(FG.Row, 3))
                End If
 
            Case InvoiceTransaction

                If Me.ExtraRetrunObject Is Nothing Then
                    Me.RetrunFrm.Retrive val(FG.TextMatrix(FG.Row, 1))
                Else
                
                    Me.ExtraRetrunObject = val(FG.TextMatrix(FG.Row, 1))
                    Me.ExtraRetrunObject1 = val(FG.TextMatrix(FG.Row, 3))
   
                    Me.ExtraRetrunObject2 = FG.TextMatrix(FG.Row, 4)
               
                 FrmReturnSalling.Retrive val(FG.TextMatrix(FG.Row, 1)) 'TEST
                End If
            
            Case Returntransaction     '  "xxx"

                If Me.ExtraRetrunObject Is Nothing Then
                   FrmReturnpurchases.Retrive val(FG.TextMatrix(FG.Row, 1))
                Else
                    Me.ExtraRetrunObject = val(FG.TextMatrix(FG.Row, 1))
                End If

            Case ShowPrice         '"xxxx"
              FrmShowPrice.Retrive val(FG.TextMatrix(FG.Row, 1))
            Case StockSettlement
              FrmStockSettlement.Retrive val(FG.TextMatrix(FG.Row, 1))
            Case InventoryOut        '”‰œ«  «·’—› «·„Œ“‰Ì
            If localindex = 0 Then
              FrmOut.Retrive val(FG.TextMatrix(FG.Row, 1))
           Else
           FrmOut1.Retrive val(FG.TextMatrix(FG.Row, 1))
           End If
            Case INVENTORYIN        '”‰œ«  «·«” ·«„  «·„Œ“‰Ì
              FrmInpout.Retrive val(FG.TextMatrix(FG.Row, 1))

            Case RowMaterialIssue        '”‰œ«  ’—› „Ê«œ Œ«„ ··«‰ «Ã
                 FrmOutProductionOrder.Retrive val(FG.TextMatrix(FG.Row, 1))

            Case ProductionMaterialReciveVoucher        '”‰œ«  «” ·«„   «‰ «Ã  «„Ã
               FrmInpoutWorkOrder.Retrive val(FG.TextMatrix(FG.Row, 1))


                     Case RowMaterialIssuesteps        '”‰œ«  ’—› „Ê«œ Œ«„ „—«Õ·  ··«‰ «Ã
                FrmOutProductionOrder1.Retrive val(FG.TextMatrix(FG.Row, 1))

               '  FrmInpoutWorkOrder1.Retrive val(FG.TextMatrix(FG.Row, 1))
             Case purchaserequest
                     If Index = 0 Then
                   FrmPO8.Retrive val(FG.TextMatrix(FG.Row, 1))
        ElseIf Index = 1 Then
         

        End If
        
                Case purchaseorder
        If Index = 0 Then
                   FrmPO5.Retrive val(FG.TextMatrix(FG.Row, 1))
        ElseIf Index = 1 Then
        FrmPO10.TxtPO6.text = FG.TextMatrix(FG.Row, 3)

        End If

     Case internalissuerequesT
         If Index = 0 Then
                     FrmPO11.Retrive val(FG.TextMatrix(FG.Row, 1))
          End If

                '«·»ÕÀ ⁄‰ «·⁄—Ê÷ «·Ã«Â“…
            Case Template
              'FrmTemplate.Retrive val(Fg.TextMatrix(Fg.Row, 1))

                '«·»ÕÀ ⁄‰ ”‰œ ’—› «·„‘«—Ì⁄
            Case Destruction
            If Index = 0 Then
                FrmDestruction.Retrive val(FG.TextMatrix(FG.Row, 1))
             ElseIf Index = 1111 Then
             FrmDestructionRet.TxtNoteSerial1.text = FG.TextMatrix(FG.Row, 3)
             'FrmDestructionRet.TXT_order_no.Text = val(Fg.TextMatrix(Fg.Row, 1))
             End If
            '«·»ÕÀ ⁄‰ „— Ã⁄ ”‰œ ’—› «·„œ›Ê⁄« 
                 Case ReturnDestruction
                FrmDestructionRet.Retrive val(FG.TextMatrix(FG.Row, 1))

                '«·»ÕÀ ⁄‰ „— Ã⁄ «·„»Ì⁄« 
            Case ReturnSalling

                If Me.ExtraRetrunObject Is Nothing Then
                 FrmReturnSalling.Retrive val(FG.TextMatrix(FG.Row, 1))
                Else
                    Me.ExtraRetrunObject = val(FG.TextMatrix(FG.Row, 1))
                End If


Case purchaseorderrequest
    FrmPO4.Retrive val(FG.TextMatrix(FG.Row, 1))
Case purchaseorder
    FrmPO5.Retrive val(FG.TextMatrix(FG.Row, 1))
   Case internalorder
      FrmPO6.Retrive val(FG.TextMatrix(FG.Row, 1))
    Case BookInventories
       FrmPO7.Retrive val(FG.TextMatrix(FG.Row, 1))
     Case purchaseOrderApproved
      FrmPO10.Retrive val(FG.TextMatrix(FG.Row, 1))
      Case salespricelistRequest
       FrmPO.Retrive val(FG.TextMatrix(FG.Row, 1))
           Case salespricelist
      FrmPO1.Retrive val(FG.TextMatrix(FG.Row, 1))
       Case SalesOrderRequest
          FrmPO2.Retrive val(FG.TextMatrix(FG.Row, 1))

                '        Case "ZZZ"
                '            FrmMoving.Retrive Val(Fg.TextMatrix(Fg.Row, 1))
            Case InsertTemplate

                If Me.FG.TextMatrix(FG.Row, FG.ColIndex("Transaction_Serial")) <> "" Then
                    DateTemp = CDate(Me.FG.TextMatrix(FG.Row, FG.ColIndex("Transaction_Serial")))

                    If DateDiff("d", Date, DateTemp) < 0 Then
                        Msg = "·ﬁœ ≈‰ ÂÌ  › —… Â–Â «·⁄—÷ ...!!!"
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Exit Sub
                    End If
                End If

               With FrmShowPrice
                  .FG.rows = 2
                    .FG.Clear flexClearScrollable, flexClearEverything
                       .FG.Refresh
                        StrSQL = "SELECT Templates.TemplateID, Template_Details.ItemID, " & "Template_Details.Quantity, Template_Details.Price, Template_Details.ItemDiscountType, " & "Template_Details.ItemDiscount, Template_Details.ItemCase, TblItems.HaveSerial " & "FROM TblItems INNER JOIN (Templates INNER JOIN Template_Details ON " & "Templates.TemplateID = Template_Details.TemplateID) ON TblItems.ItemID = " & "Template_Details.ItemID"
                      StrSQL = StrSQL + " where Templates.TemplateID=" & val(FG.TextMatrix(FG.Row, 1))
                   Set RsDetails = New ADODB.Recordset
                      RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                  .XPTxtSum.text = ""

                  If Not (RsDetails.EOF Or RsDetails.BOF) Then
                       .FG.rows = RsDetails.RecordCount + 1

                    For Num = 1 To RsDetails.RecordCount
                        .FG.TextMatrix(Num, .FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", (RsDetails("ItemID").value))
                           .FG.TextMatrix(Num, .FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
                         .FG.TextMatrix(Num, .FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
                         .FG.TextMatrix(Num, .FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
                               .FG.TextMatrix(Num, .FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
                       .FG.TextMatrix(Num, .FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
                           .FG.TextMatrix(Num, .FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))

                   If RsDetails("HaveSerial") = True Then
                            .FG.TextMatrix(Num, .FG.ColIndex("HaveSerial")) = True
                       End If

                            RsDetails.MoveNext
                      Next Num

                       End If
                       .Cala
                 End With

             Case InsertTemplateToInvoice

                 With frmsalebill
                .FG.rows = 2
               .FG.Clear flexClearScrollable, flexClearEverything
                 .FG.Refresh
                   StrSQL = "SELECT Templates.TemplateID, Template_Details.ItemID, " & "Template_Details.Quantity, Template_Details.Price, Template_Details.ItemDiscountType, " & "Template_Details.ItemDiscount, Template_Details.ItemCase, TblItems.HaveSerial " & "FROM TblItems INNER JOIN (Templates INNER JOIN Template_Details ON " & "Templates.TemplateID = Template_Details.TemplateID) ON TblItems.ItemID = " & "Template_Details.ItemID"
               StrSQL = StrSQL + " where Templates.TemplateID=" & val(FG.TextMatrix(FG.Row, 1))
                Set RsDetails = New ADODB.Recordset
                RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                 .XPTxtSum.text = ""

            If Not (RsDetails.EOF Or RsDetails.BOF) Then
                     .FG.rows = RsDetails.RecordCount + 1

                For Num = 1 To RsDetails.RecordCount
                    .FG.TextMatrix(Num, .FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", (RsDetails("ItemID").value))
                        .FG.TextMatrix(Num, .FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
                         .FG.TextMatrix(Num, .FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").value))
                          .FG.TextMatrix(Num, .FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
                      .FG.TextMatrix(Num, .FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
                        .FG.TextMatrix(Num, .FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
                          .FG.TextMatrix(Num, .FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))

                     If RsDetails("HaveSerial") = True Then
                             .FG.TextMatrix(Num, .FG.ColIndex("HaveSerial")) = True
                         End If

                            RsDetails.MoveNext
                   Next Num

                  End If

                .Cala
          End With
          
          Case InvoiceTransactionCompose

               frmsalebillCompose.Retrive val(FG.TextMatrix(FG.Row, 1))
       End Select
        End If
ElseIf Index = 1 Then

 FrmDestruction.TXT_order_no.text = val(FG.TextMatrix(FG.Row, 1))
FrmDestruction.TxtNoteSerial1.text = FG.TextMatrix(FG.Row, 3)

End If
    End If
'sak Index = 0
    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    fg_Click
    Cmd_Click (2)
End Sub

Private Sub Fg_MouseUp(Button As Integer, _
                       Shift As Integer, _
                       X As Single, _
                       Y As Single)

    If Button = vbRightButton Then
        If Shift = ShiftConstants.vbCtrlMask Then
            Me.FG.ColHidden(FG.ColIndex("Transaction_ID")) = Not Me.FG.ColHidden(FG.ColIndex("Transaction_ID"))
        End If
    End If

End Sub

Private Sub Form_Activate()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Dcombos As New ClsDataCombos
    PutFormOnTop Me.hWnd
    DCPaymentNet.Visible = False
    XPLbl(12).Visible = False
    TXTOrDer_no.Visible = False
     XPLbl(11).Visible = False
      DCEquipments.Visible = False
      lbl(62).Visible = False
    If Me.DealingForm = Returntransaction Then
        '  Me.Caption = "«·»ÕÀ ⁄‰ „— Ã⁄ «·„‘ —Ì« "
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
        '  XPLbl(1).Caption = "—ﬁ„ «·⁄„·Ì…"
        '  XPLbl(0).Caption = "«”„ «·„Ê—œ"
        '  XPChkSearchType.Caption = "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
        Dcombos.GetCustomersSuppliers 0, DCboClientsName
        Me.XPLbl(5).Visible = True
        Me.CboPayMentType.Visible = True
        '«·⁄—Ê÷ «·Ã«Â“…
    ElseIf Me.DealingForm = Template Then
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "ﬂÊœ «·⁄—÷"
        StrSQL = "SELECT * FROM Templates"
        fill_combo DCboClientsName, StrSQL
        Me.DcboStores.Visible = False
        lbl(0).Visible = False
        CmdShowMoreOptions.Enabled = False
        CboPayMentType.Visible = False
    
    ElseIf Me.DealingForm = InsertTemplate Or Me.DealingForm = InsertTemplateToInvoice Then
        '«·⁄—Ê÷ «·Ã«Â“…
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "ﬂÊœ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = "«”„ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("ClientNmae")) = " «—ÌŒ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("StorName")) = "ﬁÌ„… «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = " «—ÌŒ ≈‰ Â«¡ «·⁄—÷"
    
        XPChkSearchType.Visible = False
        TxtVal.Visible = True
        XPLbl(2).Visible = True
        XPLbl(1).Visible = False
        XPLbl(0).Visible = False
        XPLbl(3).Visible = True
        XPLbl(4).Visible = True
     ElseIf DealingForm = PurchaseTransaction Then
     XPLbl(5).Visible = True
         XPLbl(0).Caption = "«”„ «·„Ê—œ"
         XPLbl(9).Caption = "„Ê—œ ‰ﬁœÌ"
   With Me.FG
    .TextMatrix(0, .ColIndex("ClientNmae")) = "«·„Ê—œ"
    End With
     TXTOrDer_no.Visible = True
     XPLbl(11).Visible = True
     Dcombos.GetCustomersSuppliers 2, DCboClientsName
     If SystemOptions.UserInterface = ArabicInterface Then
     XPLbl(11).Caption = "«„— ‘—«¡"
     Else
     XPLbl(11).Caption = "Purchase Order"
     End If
     ElseIf DealingForm = InvoiceTransaction Then
     TXTOrDer_no.Visible = True
     XPLbl(11).Visible = True
 '     DCPaymentNet.Visible = True
 '   XPLbl(12).Visible = True
     If SystemOptions.UserInterface = ArabicInterface Then
     XPLbl(11).Caption = "«„— »Ì⁄"
     Else
     XPLbl(11).Caption = "Sell Order"
     End If
          Dcombos.GetCustomersSuppliers 1, DCboClientsName
          
        '⁄—Ê÷ «·√”⁄«—
    ElseIf Me.DealingForm = ShowPrice Or Me.DealingForm = salespricelistRequest Or Me.DealingForm = salespricelist Or Me.DealingForm = SalesOrderRequest Then
        '«·»—‰«„Ã
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = "ﬂÊœ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = " «—ÌŒ «·⁄—÷"
        '  XPLbl(1).Caption = "—ﬁ„ «·⁄—÷"
        'XPLbl(0).Caption = "«”„ «·⁄„Ì·"
        Dcombos.GetCustomersSuppliers 0, DCboClientsName
        '«· ·›Ì« 
    
    
 ElseIf Me.DealingForm = purchaseorderrequest Then
        '«·»—‰«„Ã
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = "ﬂÊœ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = " «—ÌŒ «·⁄—÷"
        '  XPLbl(1).Caption = "—ﬁ„ «·⁄—÷"
        'XPLbl(0).Caption = "«”„ «·⁄„Ì·"
        Dcombos.GetCustomersSuppliers 2, DCboClientsName, True
        '«· ·›Ì« 
  ElseIf Me.DealingForm = purchaseorder Or Me.DealingForm = BookInventories Or Me.DealingForm = purchaseOrderApproved Or Me.DealingForm = GridTransType.purchaseorder Then
        '«·»—‰«„Ã
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = "ﬂÊœ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = " «—ÌŒ «·⁄—÷"
        '  XPLbl(1).Caption = "—ﬁ„ «·⁄—÷"
        'XPLbl(0).Caption = "«”„ «·⁄„Ì·"
        Dcombos.GetCustomersSuppliers 2, DCboClientsName, True
     
    ElseIf Me.DealingForm = internalorder Then
    Dcombos.GetCustomersSuppliers 1, DCboClientsName, True
    XPLbl(0).Caption = "«”„ «·⁄„Ì·"
   With Me.FG
    .TextMatrix(0, .ColIndex("ClientNmae")) = "«·„Ê—œ «·„Ê’Ì  "
    End With
    ElseIf Me.DealingForm = Destruction Or Me.DealingForm = ReturnDestruction Then
        
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·⁄„·Ì…"
        '  Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = " «—ÌŒ «·⁄„·Ì…"
        '  XPLbl(1).Caption = "—ﬁ„ «·⁄„·Ì…"
        '  XPLbl(0).Caption = "«”„ «·„Œ“‰"
        XPChkSearchType.Visible = False
  '      StrSQL = "SELECT * From TblStore"
  '      fill_combo DCboClientsName, StrSQL
       DCboClientsName.Visible = False
XPLbl(0).Visible = False
    ElseIf Me.DealingForm = InvoiceTransactionCompose Then
          XPLbl(5).Visible = True
              XPLbl(0).Caption = "«”„ «·„Ê—œ"
              XPLbl(9).Caption = "„Ê—œ ‰ﬁœÌ"
        With Me.FG
         .TextMatrix(0, .ColIndex("ClientNmae")) = "«·„Ê—œ"
         End With
          TXTOrDer_no.Visible = True
          XPLbl(11).Visible = True
          Dcombos.GetCustomersSuppliers 2, DCboClientsName
    
    Else
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = "—ﬁ„ «·”‰œ"
        '  XPLbl(1).Caption = "—ﬁ„ «·”‰œ"
        If SystemOptions.UserInterface = ArabicInterface Then
           XPLbl(0).Caption = "««·⁄„Ì·/«·„Ê—œ"
           Else
           XPLbl(0).Caption = "«Cus/Supp"
           End If
        '  XPChkSearchType.Caption = "«”„ «·⁄„Ì· »«·ﬂ«„· ›ﬁÿ"
        Dcombos.GetCustomersSuppliers 0, DCboClientsName
        Me.XPLbl(5).Visible = True
        Me.CboPayMentType.Visible = True
    End If
    If Me.DealingForm = InventoryOut Then
      TXTOrDer_no.Visible = True
      XPLbl(11).Visible = True
      XPLbl(11).Caption = "—ﬁ„ «„— «· ‘€Ì·"
      DCEquipments.Visible = True
      lbl(62).Visible = True
      DCPaymentNet.Visible = True
      Dcombos.GetDocTypebyid Me.DCPaymentNet, 19
      XPLbl(12).Visible = True
      XPLbl(12).Caption = "‰Ê⁄ «·”‰œ"
    End If
    
    If Me.DealingForm = InsertTemplate Or Me.DealingForm = InsertTemplateToInvoice Then
        Cmd_Click (0)
    End If

    'StrSql = "SELECT * From TblCustemers where type=1"
    'fill_combo DCboCustemerName, StrSql
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DCboClientsName
    Dcombos.GetStores Me.DcboStores
    Dcombos.GetUsers Me.DcboUsers
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DcboUsers
    ChangeLang
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If FG.TextMatrix(FG.Row, FG.ColIndex("Transaction_ID")) <> "" And Me.ActiveControl Is FG Then
            fg_Click
        
        ElseIf Shift = vbCtrlMask Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As New ClsDataCombos

    Set rs = New ADODB.Recordset
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'CenterForm Me
    'FormPostion Me, GetPostion
 ' Index = 0
 If Index = 310 Then
    FG.ColHidden(FG.ColIndex("PurchOrderNo")) = True
    FG.ColHidden(FG.ColIndex("ManualNO")) = True
    FG.ColHidden(FG.ColIndex("TransSum")) = True
    FG.ColHidden(FG.ColIndex("PurchOrderNo")) = True
    FG.ColHidden(FG.ColIndex("ItemName")) = False
    FG.ColHidden(FG.ColIndex("OldNoteSerial1")) = False
    
    
 End If
    If Me.DealingForm = InvoiceTransaction Or Me.DealingForm = ReturnSalling Then
        FG.ColHidden(FG.ColIndex("OldNoteSerial1")) = False
        lbl(9).Visible = True
        txtOldNoteSerial1.Visible = True
    End If
    FG.WallPaper = BG.SearchWallpaper
    Dcombos.GetItemsNames Me.DCboItem
    Dcombos.GetEquipments DCEquipments
    Dcombos.GetPaymentType Me.DCPaymentNet
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboItem
    FG.ColFormat(FG.ColIndex("BillDate")) = "Medium Date"
If Me.DealingForm = Destruction Or Me.DealingForm = ReturnDestruction Then
CboPayMentType.Visible = False
Else
CboPayMentType.Visible = True
End If

    If SystemOptions.UserInterface = ArabicInterface Then

        With Me.CboPayMentType
            .Clear
            .AddItem "‰ﬁœÌ"
            .AddItem "«Ã·"
            .AddItem "«·ﬂ·"
            .AddItem "„ ⁄œœ"
        End With

    Else

        With Me.CboPayMentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
            .AddItem "All"
            .AddItem "Many Payed"
        End With
        
    End If

    CmdShowMoreOptions.value = False
    CmdShowMoreOptions_Click
    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DTPTo

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    CmdShowMoreOptions_Click

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    If SystemOptions.UserInterface = ArabicInterface Then
        Exit Sub
    End If
    
    lbl(1).Caption = "Results"
    XPLbl(3).Caption = "No"
    XPLbl(0).Caption = "Cust\Supp  "
    lbl(0).Caption = "Store"
    CmdShowMoreOptions.Caption = "Advanced"
    lbl(7).Caption = "User"
    XPLbl(5).Caption = "Payment"
    Fra(0).Caption = "Period"
    lbl(11).Caption = "From"
    lbl(10).Caption = "To"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    lbl(12).Caption = "Based on"
    Cmd(2).Caption = "Exit"
  XPLbl(6).Caption = "Manual No."
  XPLbl(7).Caption = "Total"
  XPLbl(8).Caption = "Net"
  XPLbl(12).Caption = "Type Payment"
    With Me.CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Credit"
        .AddItem "All"
         .AddItem "Many"
    End With
         XPLbl(0).Caption = "Supplier"
         XPLbl(9).Caption = "Cash Supp"
    With Me.FG

        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Vchr no"
        .TextMatrix(0, .ColIndex("BillDate")) = "Date"
        .TextMatrix(0, .ColIndex("ClientNmae")) = "Customer Name"
        .TextMatrix(0, .ColIndex("StorName")) = "Store Name"
        .TextMatrix(0, .ColIndex("BillDate1")) = "Date"
        .TextMatrix(0, .ColIndex("ManualNO")) = "Manual No."
        .TextMatrix(0, .ColIndex("TransSum")) = "Total"
        .TextMatrix(0, .ColIndex("TransNet")) = "Net Value"
        .TextMatrix(0, .ColIndex("CashCustomerPhone")) = "Phone"
        .TextMatrix(0, .ColIndex("PurchOrderNo")) = "Purch OrderNo"

    End With
 Command1.Caption = "Move Search To Screen"
    XPLbl(10).Caption = "Tel"
         
   Dim Dcombos As New ClsDataCombos

    Dcombos.GetItemsNames Me.DCboItem
    Dcombos.GetEquipments DCEquipments
    Dcombos.GetPaymentType Me.DCPaymentNet
 
 
 
     If Me.DealingForm = Returntransaction Then
        '  Me.Caption = "«·»ÕÀ ⁄‰ „— Ã⁄ «·„‘ —Ì« "
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
        '  XPLbl(1).Caption = "—ﬁ„ «·⁄„·Ì…"
        '  XPLbl(0).Caption = "«”„ «·„Ê—œ"
        '  XPChkSearchType.Caption = "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
        Dcombos.GetCustomersSuppliers 0, DCboClientsName
        Me.XPLbl(5).Visible = True
        Me.CboPayMentType.Visible = True
        '«·⁄—Ê÷ «·Ã«Â“…
    ElseIf Me.DealingForm = Template Then
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "ﬂÊœ «·⁄—÷"
        Dim StrSQL  As String
        StrSQL = "SELECT * FROM Templates"
        fill_combo DCboClientsName, StrSQL
        Me.DcboStores.Visible = False
        lbl(0).Visible = False
        CmdShowMoreOptions.Enabled = False
        CboPayMentType.Visible = False
    
    ElseIf Me.DealingForm = InsertTemplate Or Me.DealingForm = InsertTemplateToInvoice Then
        '«·⁄—Ê÷ «·Ã«Â“…
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "ﬂÊœ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = "«”„ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("ClientNmae")) = " «—ÌŒ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("StorName")) = "ﬁÌ„… «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = " «—ÌŒ ≈‰ Â«¡ «·⁄—÷"
    
        XPChkSearchType.Visible = False
        TxtVal.Visible = True
        XPLbl(2).Visible = True
        XPLbl(1).Visible = False
        XPLbl(0).Visible = False
        XPLbl(3).Visible = True
        XPLbl(4).Visible = True
     ElseIf DealingForm = PurchaseTransaction Then
     XPLbl(5).Visible = True
         XPLbl(0).Caption = "Supplier"
         XPLbl(9).Caption = "Cash Supp"
   With Me.FG
    .TextMatrix(0, .ColIndex("ClientNmae")) = "Supplier"
    End With
     TXTOrDer_no.Visible = True
     XPLbl(11).Visible = True
     Dcombos.GetCustomersSuppliers 2, DCboClientsName
     If SystemOptions.UserInterface = ArabicInterface Then
     XPLbl(11).Caption = "«„— ‘—«¡"
     Else
     XPLbl(11).Caption = "Purchase Order"
     End If
     ElseIf DealingForm = InvoiceTransaction Then
     TXTOrDer_no.Visible = True
     XPLbl(11).Visible = True
 '     DCPaymentNet.Visible = True
 '   XPLbl(12).Visible = True
     If SystemOptions.UserInterface = ArabicInterface Then
     XPLbl(11).Caption = "«„— »Ì⁄"
     Else
     XPLbl(11).Caption = "Sell Order"
     End If
          Dcombos.GetCustomersSuppliers 1, DCboClientsName
          
        '⁄—Ê÷ «·√”⁄«—
    ElseIf Me.DealingForm = ShowPrice Or Me.DealingForm = salespricelistRequest Or Me.DealingForm = salespricelist Or Me.DealingForm = SalesOrderRequest Then
        '«·»—‰«„Ã
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = "ﬂÊœ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = " «—ÌŒ «·⁄—÷"
        '  XPLbl(1).Caption = "—ﬁ„ «·⁄—÷"
        'XPLbl(0).Caption = "«”„ «·⁄„Ì·"
        Dcombos.GetCustomersSuppliers 0, DCboClientsName
        '«· ·›Ì« 
    
    
 ElseIf Me.DealingForm = purchaseorderrequest Then
        '«·»—‰«„Ã
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = "ﬂÊœ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = " «—ÌŒ «·⁄—÷"
        '  XPLbl(1).Caption = "—ﬁ„ «·⁄—÷"
        'XPLbl(0).Caption = "«”„ «·⁄„Ì·"
        Dcombos.GetCustomersSuppliers 2, DCboClientsName, True
        '«· ·›Ì« 
  ElseIf Me.DealingForm = purchaseorder Or Me.DealingForm = BookInventories Or Me.DealingForm = purchaseOrderApproved Or Me.DealingForm = GridTransType.purchaseorder Then
        '«·»—‰«„Ã
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = "ﬂÊœ «·⁄—÷"
        '  Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = " «—ÌŒ «·⁄—÷"
        '  XPLbl(1).Caption = "—ﬁ„ «·⁄—÷"
        'XPLbl(0).Caption = "«”„ «·⁄„Ì·"
        Dcombos.GetCustomersSuppliers 2, DCboClientsName, True
     
    ElseIf Me.DealingForm = internalorder Then
    Dcombos.GetCustomersSuppliers 1, DCboClientsName, True
    XPLbl(0).Caption = "Customer"
    XPLbl(0).Caption = "Supp\Cust"
   With Me.FG
    .TextMatrix(0, .ColIndex("ClientNmae")) = "Supplier Reco "
    End With
    ElseIf Me.DealingForm = Destruction Or Me.DealingForm = ReturnDestruction Then
        
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·⁄„·Ì…"
        '  Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = " «—ÌŒ «·⁄„·Ì…"
        '  XPLbl(1).Caption = "—ﬁ„ «·⁄„·Ì…"
        '  XPLbl(0).Caption = "«”„ «·„Œ“‰"
        XPChkSearchType.Visible = False
  '      StrSQL = "SELECT * From TblStore"
  '      fill_combo DCboClientsName, StrSQL
       DCboClientsName.Visible = False
XPLbl(0).Visible = False
    ElseIf Me.DealingForm = InvoiceTransactionCompose Then
          XPLbl(5).Visible = True
              XPLbl(0).Caption = "Supplier"
              XPLbl(9).Caption = "Cach Supp"
        With Me.FG
         .TextMatrix(0, .ColIndex("ClientNmae")) = "Supplier"
         End With
          TXTOrDer_no.Visible = True
          XPLbl(11).Visible = True
          Dcombos.GetCustomersSuppliers 2, DCboClientsName
    
    Else
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
        '  Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = "—ﬁ„ «·”‰œ"
        '  XPLbl(1).Caption = "—ﬁ„ «·”‰œ"
        
           XPLbl(0).Caption = "«Cus/Supp"
       
        '  XPChkSearchType.Caption = "«”„ «·⁄„Ì· »«·ﬂ«„· ›ﬁÿ"
        Dcombos.GetCustomersSuppliers 0, DCboClientsName
        Me.XPLbl(5).Visible = True
        Me.CboPayMentType.Visible = True
    End If
    If Me.DealingForm = InventoryOut Then
      TXTOrDer_no.Visible = True
      XPLbl(11).Visible = True
      XPLbl(11).Caption = "Batch Order"
      DCEquipments.Visible = True
      lbl(62).Visible = True
      DCPaymentNet.Visible = True
      Dcombos.GetDocTypebyid Me.DCPaymentNet, 19
      
      XPLbl(12).Caption = "Type"
    End If
    
  lbl(137).Caption = "Purchase Order"
    Dcombos.GetStores Me.DcboStores
    Dcombos.GetUsers Me.DcboUsers
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
allTransaction_ID = ""
    If Me.DealingForm = InvoiceTransaction Or Me.DealingForm = PurchaseTransaction Or Me.DealingForm = InventoryOut Then

        'Set Me.FG.DataSource = rs
        Dim Transaction_Type As Double
        If Not (rs.EOF Or rs.BOF) Then
            FG.rows = rs.RecordCount + 1

            For Num = 1 To rs.RecordCount

                With FG
                    .TextMatrix(Num, .ColIndex("Count")) = Num
                  .TextMatrix(Num, .ColIndex("ContainerNo")) = rs!ContainerNo & ""
                    
                    .TextMatrix(Num, .ColIndex("CarPrevValue")) = IIf(IsNull(rs("CarPrevValue").value), "", (rs("CarPrevValue").value))
                    .TextMatrix(Num, .ColIndex("PlateNo")) = IIf(IsNull(rs("PlateNo").value), "", (rs("PlateNo").value))
                  .TextMatrix(Num, .ColIndex("PurchOrderNo")) = IIf(IsNull(rs("order_no").value), "", (rs("order_no").value))
                    .TextMatrix(Num, .ColIndex("OldNoteSerial1")) = IIf(IsNull(rs("OldNoteSerial1").value), "", (rs("OldNoteSerial1").value))
                    .TextMatrix(Num, .ColIndex("noteserial1")) = IIf(IsNull(rs("noteserial1").value), "", (rs("noteserial1").value))
                    .TextMatrix(Num, .ColIndex("ManualNO")) = IIf(IsNull(rs("ManualNO").value), "", (rs("ManualNO").value))
                    If Index = 310 Then
                    .TextMatrix(Num, .ColIndex("ItemId")) = IIf(IsNull(rs("ItemId").value), "", (rs("ItemId").value))
                    .TextMatrix(Num, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", (rs("ItemName").value))
                    End If
                    .TextMatrix(Num, .ColIndex("Fullcode")) = IIf(IsNull(rs("Fullcode").value), "", (rs("Fullcode").value))
                    .TextMatrix(Num, .ColIndex("Address")) = IIf(IsNull(rs("Address").value), "", (rs("Address").value))
                    .TextMatrix(Num, .ColIndex("Cus_mobile")) = IIf(IsNull(rs("Cus_mobile").value), "", (rs("Cus_mobile").value))
                    
                    
                If val(TxtTotalValue.text) <> 0 Or val(TxtNetValue.text) <> 0 Then
                    .TextMatrix(Num, .ColIndex("TransSum")) = IIf(IsNull(rs("TransSum").value), 0, (rs("TransSum").value))
                    .TextMatrix(Num, .ColIndex("TransNet")) = IIf(IsNull(rs("TransNet").value), 0, (rs("TransNet").value))
                    End If
                    
                    .TextMatrix(Num, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
                    .TextMatrix(Num, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
                    allTransaction_ID = allTransaction_ID & "," & .TextMatrix(Num, .ColIndex("Transaction_ID"))
                   .TextMatrix(Num, .ColIndex("CashCustomerPhone")) = IIf(IsNull(rs("CashCustomerPhone").value), "", (rs("CashCustomerPhone").value))
                    If Not IsNull(rs("Transaction_Date").value) Then
                     '   .TextMatrix(Num, .ColIndex("BillDate")) = Format(rs("Transaction_Date").value, "dd/mm/yyyy")
                        .TextMatrix(Num, .ColIndex("BillDate")) = (rs("Transaction_Date").value)
                          .TextMatrix(Num, .ColIndex("BillDate1")) = (rs("Transaction_Date").value)
                        
                    Else
                        .TextMatrix(Num, .ColIndex("BillDate")) = ""
                    End If
                 '   .TextMatrix(Num, .ColIndex("BillDate1")) = (rs("Transaction_Date").value)
                    

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(rs("CashCustomerName").value), IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value)), Trim(rs("CashCustomerName").value))
                        .TextMatrix(Num, .ColIndex("StorName")) = IIf(IsNull(rs("StoreName").value), "", Trim(rs("StoreName").value))
                    Else
                        .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(rs("CashCustomerName").value), IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value)), Trim(rs("CashCustomerName").value))
                        .TextMatrix(Num, .ColIndex("StorName")) = IIf(IsNull(rs("StoreNamee").value), "", Trim(rs("StoreNamee").value))
                    End If



                End With

                rs.MoveNext
            Next Num

        End If
                
         FG.AutoSize 0, FG.Cols - 1, False
    ElseIf Me.DealingForm = Template Or Me.DealingForm = InsertTemplate Or Me.DealingForm = InsertTemplateToInvoice Then

        If Not (rs.EOF Or rs.BOF) Then
            FG.rows = rs.RecordCount + 1

            For Num = 1 To rs.RecordCount

                With FG
                    .TextMatrix(Num, .ColIndex("Count")) = Num
                
                    .TextMatrix(Num, .ColIndex("noteserial1")) = IIf(IsNull(rs("noteserial1").value), "", (rs("noteserial1").value))
                    .TextMatrix(Num, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("TemplateID").value), "", val(rs("TemplateID").value))
                    .TextMatrix(Num, .ColIndex("BillDate")) = IIf(IsNull(rs("TemplateName").value), "", (rs("TemplateName").value))
                    .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(rs("Date").value), "", Format(rs("Date").value, "yyyy/m/d"))
                    .TextMatrix(Num, .ColIndex("StorName")) = IIf(IsNull(rs("Summition").value), "", Trim(rs("Summition").value))

                    If Not IsNull(rs("TemplateTime").value) Then
                        .TextMatrix(Num, .ColIndex("Transaction_Serial")) = (rs("TemplateTime").value)
                    End If

                End With

                rs.MoveNext
            Next Num

        End If

    ElseIf Not (rs.EOF Or rs.BOF) Then
        FG.rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("Count")) = Num
                If Me.DealingForm = ReturnDestruction Then
                 .TextMatrix(Num, .ColIndex("noteserial1")) = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
                 Else
                .TextMatrix(Num, .ColIndex("noteserial1")) = IIf(IsNull(rs("noteserial1").value), "", (rs("noteserial1").value))
                .TextMatrix(Num, .ColIndex("OldNoteSerial1")) = IIf(IsNull(rs("OldNoteSerial1").value), "", (rs("OldNoteSerial1").value))
               End If
                .TextMatrix(Num, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
                .TextMatrix(Num, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
Transaction_Type = IIf(IsNull(rs("Transaction_Type").value), "", (rs("Transaction_Type").value))

                  
                If Not IsNull(rs("Transaction_Date").value) Then
                    .TextMatrix(Num, .ColIndex("BillDate")) = rs("Transaction_Date").value
                      .TextMatrix(Num, .ColIndex("BillDate1")) = (rs("Transaction_Date").value)
                Else
                    .TextMatrix(Num, .ColIndex("BillDate")) = ""
                End If

             '   .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
             '   .TextMatrix(Num, .ColIndex("StorName")) = IIf(IsNull(rs("StoreName").value), "", Trim(rs("StoreName").value))
            

If Transaction_Type = 46 Then

     .TextMatrix(Num, .ColIndex("BillDate1")) = .TextMatrix(Num, .ColIndex("BillDate1")) & " -  ⁄—÷ ”⁄—"
ElseIf Transaction_Type = 38 Then

     .TextMatrix(Num, .ColIndex("BillDate1")) = .TextMatrix(Num, .ColIndex("BillDate1")) & " -  ÿ·» œ«Œ·Ì"
                            
                  End If
                  
                         If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                        .TextMatrix(Num, .ColIndex("StorName")) = IIf(IsNull(rs("StoreName").value), "", Trim(rs("StoreName").value))
                    Else
                        .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                        .TextMatrix(Num, .ColIndex("StorName")) = IIf(IsNull(rs("StoreNamee").value), "", Trim(rs("StoreNamee").value))
                    End If
            End With

            rs.MoveNext
        Next Num

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    FG.SetFocus
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    FormPostion Me, SavePostion
    Set Me.ExtraRetrunObject = Nothing

    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql() As String
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim MySQL As String
    Dim m_SearchFrom As GridTransType
    Dim Begin As Boolean
    Dim StrWhere As String

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        MySQL = "SELECT DISTINCT Transactions.Transaction_ID,Transactions.OldNoteSerial1,Transactions.order_no,Transactions.ContainerNo, Transactions.PurchOrderNo,Transactions.Transaction_Serial," & "Transactions.Transaction_Date,TblCustemers.CusName, TblStore.StoreName ,  dbo.Transactions.ManualNO"
        MySQL = MySQL + " FROM (TblStore RIGHT JOIN (TblCustemers RIGHT JOIN Transactions " & "ON TblCustemers.CusID=Transactions.CusID) ON TblStore.StoreID=Transactions.StoreID)" & "INNER JOIN Transaction_Details ON Transactions.Transaction_ID=Transaction_Details.Transaction_ID "
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
    '    MySQL = "SELECT DISTINCT Transactions.Transaction_ID,Transactions.Transaction_Serial," & "convert(nvarchar(50),Transactions.Transaction_Date,111)as Transaction_Date,TblCustemers.CusName, TblStore.StoreNamee ,TblCustemers.CusNamee, TblStore.StoreName,Transactions.NoteSerial1"
    '    MySQL = MySQL + " FROM (TblStore RIGHT JOIN (TblCustemers RIGHT JOIN Transactions " & "ON TblCustemers.CusID=Transactions.CusID) ON TblStore.StoreID=Transactions.StoreID)" & "INNER JOIN Transaction_Details ON Transactions.Transaction_ID=Transaction_Details.Transaction_ID "
   ' End If
   
   
   If val(TxtTotalValue.text) <> 0 Or val(TxtNetValue.text) <> 0 Then
  MySQL = " SELECT  DISTINCT    dbo.Transactions.Transaction_ID,Transactions.OldNoteSerial1,Transactions.ContainerNo, dbo.Transactions.Transaction_Serial, CONVERT(nvarchar(50), dbo.Transactions.Transaction_Date, 111) AS Transaction_Date,"
  MySQL = MySQL + "                    dbo.TblCustemers.CusName, dbo.TblStore.StoreNamee,Transactions.PlateNo, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.Transactions.NoteSerial1,"
  MySQL = MySQL + "                    QryTransactionsTotal.TransSum , Transactions.CarPrevValue,Transactions.PlateNo, QryTransactionsTotal.TransNet , dbo.Transactions.Transaction_Type, dbo.Transactions.ManualNO , dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone ,Transactions.order_no,TblCustemers.CusID"
  MySQL = MySQL + "  FROM         dbo.TblStore RIGHT OUTER JOIN"
  MySQL = MySQL + "                     dbo.TblCustemers RIGHT OUTER JOIN"
  MySQL = MySQL + "                     dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
  MySQL = MySQL + "                    dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
  MySQL = MySQL + "                    dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID"
  Else
  MySQL = " SELECT DISTINCT"
  MySQL = MySQL + "                     dbo.Transactions.Transaction_ID,Transactions.OldNoteSerial1, Transactions.CarPrevValue,Transactions.PlateNo,Transactions.ContainerNo, dbo.Transactions.Transaction_Serial, CONVERT(nvarchar(50), dbo.Transactions.Transaction_Date, 111) AS Transaction_Date,"
  MySQL = MySQL + "                     dbo.TblCustemers.CusName,TblCustemers.FullCode,TblCustemers.Address, TblCustemers.Cus_mobile,Transactions.order_no,Transactions.PurchOrderNo,dbo.TblStore.StoreNamee, dbo.TblCustemers.CusNamee, dbo.TblStore.StoreName, dbo.Transactions.NoteSerial1,"
  MySQL = MySQL + "                     dbo.Transactions.Transaction_Type , dbo.Transactions.ManualNO , dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone ,Transactions.order_no,TblCustemers.CusID"
  If Index = 310 Then
        MySQL = MySQL + "                     ,dbo.Transactions.Transaction_Type ,tblItems.ItemName,tblItems.ItemNamee,tblItems.ItemId "
  End If
  MySQL = MySQL + "  FROM         dbo.TblStore RIGHT OUTER JOIN"
  MySQL = MySQL + "                     dbo.TblCustemers RIGHT OUTER JOIN"
    If Me.DealingForm <> InvoiceTransactionCompose Then
        MySQL = MySQL + "                     dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID ON dbo.TblStore.StoreID = dbo.Transactions.StoreID left outer JOIN "
    ElseIf Me.DealingForm = InvoiceTransactionCompose Then
        MySQL = MySQL + "                     dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.FarmID ON dbo.TblStore.StoreID = dbo.Transactions.StoreID left outer JOIN "
    End If
  MySQL = MySQL + "                     dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
  MySQL = MySQL + "                     Left outer join dbo.tblItems  ON dbo.tblItems.ItemID = dbo.Transaction_Details.Item_ID "
 
  End If
  
End If
    m_SearchFrom = Me.DealingForm

    Select Case m_SearchFrom

        Case PurchaseTransaction
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=1 or dbo.Transactions.Transaction_Type=22"
            Begin = True

        Case InvoiceTransaction
            StrSQL = MySQL + " WHERE (dbo.Transactions.Transaction_Type=2 or dbo.Transactions.Transaction_Type=21)"
                    If localindex = 4 Then
                        StrSQL = StrSQL + " and IsNull(Transactions.TypeInvoice,0) = 1"
                    ElseIf localindex = 5 Then
                            StrSQL = StrSQL + " and IsNull(Transactions.TypeInvoice,0) = 2"
                 End If
            If val(CboPayMentType.ListIndex) = 3 And val(DCPaymentNet.BoundText) <> 0 Then
            StrSQL = StrSQL & " and ( dbo.transactions.Transaction_ID in (select TransID from  TblSalesPayment where PaymentID=" & val(DCPaymentNet.BoundText) & " and Value>0  ) )"
            End If
            If Index = 131 Then
            Dim intDef As Integer
            
            StrSQL = StrSQL & " and dbo.Transactions.StoreID=" & val(frmsalebill2.DCboStoreName.BoundText) & " "
            End If
            Begin = True

Case StockSettlement
           StrSQL = MySQL + " WHERE (dbo.Transactions.Transaction_Type=15 or dbo.Transactions.Transaction_Type=16) "
            Begin = True


        Case Returntransaction
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=5"
            Begin = True
        Case 666
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=6 and IsNull(chkIsFirstInv,0) = 1 "
            Begin = True

        Case ShowPrice
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=17"
            Begin = True
    Case salespricelistRequest
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=41"
            Begin = True
        Case salespricelist
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=42"
            Begin = True
    
   Case SalesOrderRequest
             StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=43"
            Begin = True
            
        Case InventoryOut
        If localindex = 0 Then
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=19"
        Else
        StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=36"
        End If
        
            Begin = True
        
        Case INVENTORYIN
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=20"
            Begin = True
        
            If Index = 10 Then
             StrSQL = MySQL + " and CBoBasedON=11 "
            End If
            
        
        Case RowMaterialIssue
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=27"
            Begin = True
        
        Case RowMaterialIssuesteps
        
             StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=33"
            Begin = True
       
       
        Case ProductionMaterialReciveVoucher
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=28"
            Begin = True
        
        Case GridTransType.purchaserequest
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=47"
            Begin = True
        If Index <> 0 Then
        If CheckAprroveScreen("FrmPO8") = True Then
StrSQL = StrSQL + " and approved=1"
End If
End If
         Case ProductionMaterialReciveVoucherStEPS
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=34"
            Begin = True
        
        
            '”‰œ ’—› «·„‘«—Ì⁄
        Case Destruction
        
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=18"
            
            Begin = True
      Case ReturnDestruction
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=66"
            
            Begin = True
            
        Case ReturnSalling
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=9"
            Begin = True
Case purchaseorderrequest
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=45"
            Begin = True
 Case purchaseorder
 
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=46"
            Begin = True
 Case internalissuerequesT
 
  StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=62"
            Begin = True
            
            Case internalorder


If Me.Index = 6 Then
'»ÕÀ ÿ·» œ«Œ·Ì Ê ÿ·» ‘—«¡
      StrSQL = MySQL + " WHERE ( dbo.Transactions.Transaction_Type=46 or  dbo.Transactions.Transaction_Type=38)"
            Begin = True
            
            
If CheckAprroveScreen("FrmPO5") = True Then
StrSQL = StrSQL + " and approved=1"
End If


Else

'»ÕÀ ÿ·» œ«Œ·Ì ›ﬁÿ

              
              
                     StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=38"
              
              If SystemOptions.IsInternalMultiOrder Then
                    StrSQL = StrSQL & " AND Transactions.Transaction_ID NOT IN ("
    
                    StrSQL = StrSQL & " SELECT"
      
                    StrSQL = StrSQL & " OrderID"
       
                    StrSQL = StrSQL & " FROM   Transaction_Details AS td"
                    StrSQL = StrSQL & " INNER JOIN Transactions t"
                    StrSQL = StrSQL & " ON  t.Transaction_ID = td.Transaction_ID"
                    StrSQL = StrSQL & " AND t.Transaction_Type = 10"
                    StrSQL = StrSQL & " AND t.BillBasedOn = 1"
                    StrSQL = StrSQL & " INNER JOIN Transaction_Details tt"
                    StrSQL = StrSQL & " ON  tt.Transaction_ID = t.OrderID"
                    StrSQL = StrSQL & " AND tt.Item_ID = td.Item_ID"
                    StrSQL = StrSQL & " AND tt.UnitId = td.UnitId"
                    StrSQL = StrSQL & " Group By"
                    StrSQL = StrSQL & " td.Item_ID,"
                    StrSQL = StrSQL & " td.UnitId,tt.ShowQty,"
                    StrSQL = StrSQL & " OrderID"
                    StrSQL = StrSQL & " Having SUM(td.ShowQty) >= (tt.ShowQty)"

                    StrSQL = StrSQL & " )"
            End If
     
            Begin = True
            
   If Index <> 0 Then
If CheckAprroveScreen("FrmPO6") = True Then
StrSQL = StrSQL + " and approved=1"
End If
End If

End If




            
  Case internalissuerequesT
  StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=62"
            Begin = True
Case InvoiceTransactionCompose
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=40"
 
            If Index = 6 Then
            StrSQL = StrSQL + " and OrderType=3 "
          '   StrSQL = StrSQL + " and NoteSerial1 not in ( "
             
          '       StrSQL = StrSQL + "  SELECT DISTINCT NotSeialPO6"
    'StrSQL = StrSQL + " From dbo.Transactions"
    'StrSQL = StrSQL + "  WHERE     (dbo.Transactions.Transaction_Typ = 29)"

    '         StrSQL = StrSQL + ")"
            
            End If

            
            Begin = True
Case BookInventories
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=39"
            Begin = True
Case purchaseOrderApproved
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=29"
            Begin = True


            
        '    StrSQL = MySQL + " WHERE BranchId=29"
            
            '    Case "ZZZ"  '«· ÕÊÌ· „‰ „Œ“‰ ≈·Ï „Œ“‰
            '        StrSql = "select * From QRyTransSearch WHERE dbo.Transactions.Transaction_Type=10"
            '«·⁄—Ê÷ «·Ã«Â“…
        Case Template, InsertTemplate, InsertTemplateToInvoice

            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = "select * From TemplateSearch"
                Begin = False
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = "SELECT TemplateSearch.* FROM dbo.TemplateSearch() TemplateSearch"
                Begin = False
            End If

        Case ProductionOrder
            StrSQL = MySQL + " WHERE dbo.Transactions.Transaction_Type=26"
            Begin = True
    End Select

    
    
    If m_SearchFrom = Template Or m_SearchFrom = InsertTemplate Or m_SearchFrom = InsertTemplateToInvoice Then
        If XPTxtBillNum.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and TemplateID like '%" & (XPTxtBillNum.text) & "%'"
            Else
                StrWhere = StrWhere + " where TemplateID like '%" & (XPTxtBillNum.text) & "%'"
                Begin = True
            End If
        End If



  If Trim(txtContainerNo.text) <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.ContainerNo ='" & Trim(txtContainerNo.text) & "'"
            Else
                StrWhere = StrWhere + " where Transactions.ContainerNo ='" & Trim(txtContainerNo.text) & "'"
                Begin = True
            End If
        End If
        
        
  If val(txtPrevValue.text) <> 0 Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.CarPrevValue =" & val(txtPrevValue.text) & ""
            Else
                StrWhere = StrWhere + " where Transactions.CarPrevValue =" & val(txtPrevValue.text) & ""
                Begin = True
            End If
        End If
        
        
       

  If Trim(TxtPlatNo.text) <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.PlateNo ='" & Trim(TxtPlatNo.text) & "'"
            Else
                StrWhere = StrWhere + " where Transactions.PlateNo ='" & Trim(TxtPlatNo.text) & "'"
                Begin = True
            End If
        End If
         
        
        
        If DCboClientsName.BoundText <> "" And DCboClientsName.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and TemplateID =" & Trim(DCboClientsName.BoundText)
            Else
                StrWhere = StrWhere + " where TemplateID =" & Trim(DCboClientsName.BoundText)
                Begin = True
            End If
        End If

        If TxtVal.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and Summition =" & TxtVal.text
            Else
                StrWhere = StrWhere + " where Summition=" & TxtVal.text
                Begin = True
            End If
        End If

        
        If Not IsNull(Me.DTPFrom.value) Then
            If Begin = True Then
                StrWhere = StrWhere + " and [Date] >=" & SQLDate(Me.DTPFrom.value, True) & ""
            Else
                StrWhere = StrWhere + " where [Date] >=" & SQLDate(Me.DTPFrom.value, True) & ""
                Begin = True
            End If
        End If

        If Not IsNull(Me.DTPTo.value) Then
            If Begin = True Then
                StrWhere = StrWhere + " and [Date] <=" & SQLDate(Me.DTPTo.value, True) & ""
            Else
                StrWhere = StrWhere + " where [Date] <=" & SQLDate(Me.DTPTo.value, True) & ""
                Begin = True
            End If
        End If

        Build_Sql = StrSQL + StrWhere + " order by TemplateID"
    ElseIf m_SearchFrom = Destruction Or m_SearchFrom = ReturnDestruction Then '«· ·›Ì« 

        If XPTxtBillNum.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and NoteSerial1 like '%" & (XPTxtBillNum.text) & "%'"
            Else
                StrWhere = StrWhere + " where NoteSerial1 like '%" & (XPTxtBillNum.text) & "%'"
                Begin = True
            End If
        End If

        If DCboClientsName.BoundText <> "" And DCboClientsName.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and StoreID =" & Trim(DCboClientsName.BoundText)
            Else
                StrWhere = StrWhere + " where StoreID =" & Trim(DCboClientsName.BoundText)
                Begin = True
            End If
        End If


      If SystemOptions.usertype <> UserAdminAll Then
            If m_SearchFrom = 0 And localindex = 4 Then
            Else
                StrWhere = StrWhere & " AND   Transactions.BranchId=" & Current_branch
            End If
            End If
            
            
  '*************************************************************************************
       If Not IsNull(Me.DTPFrom.value) Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.Transaction_date >=" & SQLDate(Me.DTPFrom.value, True) & ""
            Else
                StrWhere = StrWhere + " where Transactions.Transaction_date >=" & SQLDate(Me.DTPFrom.value, True) & ""
                Begin = True
            End If
        End If

        If Not IsNull(Me.DTPTo.value) Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.Transaction_date <=" & SQLDate(Me.DTPTo.value, True) & ""
            Else
                StrWhere = StrWhere + " where Transactions.Transaction_date <=" & SQLDate(Me.DTPTo.value, True) & ""
                Begin = True
            End If
        End If

        If Me.DcboStores.BoundText <> "" And DcboStores.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.StoreID=" & Me.DcboStores.BoundText & ""
            Else
                StrWhere = StrWhere + " where Transactions.StoreID=" & Me.DcboStores.BoundText & ""
                Begin = True
            End If
        End If

        If Me.DCboItem.text <> "" Then

            'If Me.DCboItem.BoundText <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transaction_Details.Item_ID=" & Me.DCboItem.BoundText & ""
            Else
                StrWhere = StrWhere + " where Transaction_Details.Item_ID=" & Me.DCboItem.BoundText & ""
                Begin = True
            End If
        End If

        If Trim(txtPurchOrderNo.text) <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and transactions.order_no=N'" & (txtPurchOrderNo.text) & "'"
            Else
                StrWhere = StrWhere + " where transactions.order_no=N'" & Trim(txtPurchOrderNo.text) & "'"
                Begin = True
            End If
        End If


    If Index = 310 Then
        MySQL = MySQL + "                     and  dbo.Transaction_Details.Item_ID = " & mmItemId
    End If



        If val(TxtItemQty.text) > 0 Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transaction_Details.Quantity=" & val(TxtItemQty.text) & ""
            Else
                StrWhere = StrWhere + " where Transaction_Details.Quantity=" & val(TxtItemQty.text) & ""
                Begin = True
            End If
        End If

        If val(TxtItemPrice.text) > 0 Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transaction_Details.Price=" & val(TxtItemPrice.text) & ""
            Else
                StrWhere = StrWhere + " where Transaction_Details.Price=" & val(TxtItemPrice.text) & ""
                Begin = True
            End If
        End If

        If Trim(Me.TxtItemSerial.text) <> "" Then
            If ChkSerialSearchType.value = vbChecked Then
                If Begin = True Then
                    StrWhere = StrWhere + " and Transaction_Details.ItemSerial='" & Trim(TxtItemSerial.text) & "'"
                Else
                    StrWhere = StrWhere + " where Transaction_Details.ItemSerial='" & Trim(TxtItemSerial.text) & "'"
                    Begin = True
                End If

            ElseIf ChkSerialSearchType.value = vbUnchecked Then

                If Begin = True Then
                    StrWhere = StrWhere + " and Transaction_Details.ItemSerial like '%" & Trim(TxtItemSerial.text) & "%'"
                Else
                    StrWhere = StrWhere + " where Transaction_Details.ItemSerial like '%" & Trim(TxtItemSerial.text) & "%'"
                    Begin = True
                End If
            End If
        End If

  '*************************************************************************************
     If Not SystemOptions.IsHiddenUser Then
        StrWhere = StrWhere & " and IsNull(Transactions.IsHiddenInv,0) =0"
     
     End If

     If SystemOptions.usertype <> UserAdminAll Then
 
          If SystemOptions.FixedCustomer = 1 Then
            If m_SearchFrom = 0 And localindex = 4 Then
            Else
                StrWhere = StrWhere & " and    Transactions.UserID = " & user_id
            End If
             End If
  
  End If
        Build_Sql = StrSQL + StrWhere + " Order by Transactions.Transaction_ID"
    Else

        '---------------------------------
        
           If val(TxtTotalValue.text) <> 0 Then
            If Begin = True Then
            If RdTotal(0).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransSum > " & val(TxtTotalValue.text) & ""
            ElseIf RdTotal(1).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransSum < " & val(TxtTotalValue.text) & ""
            ElseIf RdTotal(2).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransSum = " & val(TxtTotalValue.text) & ""
            ElseIf RdTotal(3).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransSum >= " & val(TxtTotalValue.text) & ""
            ElseIf RdTotal(4).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransSum <= " & val(TxtTotalValue.text) & ""
            End If
            Else
            If RdTotal(0).value = True Then
                 StrWhere = StrWhere + " and QryTransactionsTotal.TransSum > " & val(TxtTotalValue.text) & ""
            ElseIf RdTotal(1).value = True Then
                 StrWhere = StrWhere + " and QryTransactionsTotal.TransSum < " & val(TxtTotalValue.text) & ""
            ElseIf RdTotal(2).value = True Then
                 StrWhere = StrWhere + " and QryTransactionsTotal.TransSum = " & val(TxtTotalValue.text) & ""
            ElseIf RdTotal(3).value = True Then
                 StrWhere = StrWhere + " and QryTransactionsTotal.TransSum >= " & val(TxtTotalValue.text) & ""
            ElseIf RdTotal(4).value = True Then
                 StrWhere = StrWhere + " and QryTransactionsTotal.TransSum <= " & val(TxtTotalValue.text) & ""
            End If
            
                Begin = True
            End If
        End If
       If val(TxtNetValue.text) <> 0 Then
            If Begin = True Then
            If RdNet(0).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransNet  > " & val(TxtNetValue.text) & ""
            ElseIf RdNet(1).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransNet  < " & val(TxtNetValue.text) & ""
            ElseIf RdNet(2).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransNet  = " & val(TxtNetValue.text) & ""
            ElseIf RdNet(3).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransNet  >= " & val(TxtNetValue.text) & ""
           ElseIf RdNet(4).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransNet  <= " & val(TxtNetValue.text) & ""
            End If
            Else
            If RdNet(0).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransNet  > " & val(TxtNetValue.text) & ""
            ElseIf RdNet(1).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransNet  < " & val(TxtNetValue.text) & ""
            ElseIf RdNet(2).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransNet  = " & val(TxtNetValue.text) & ""
            ElseIf RdNet(3).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransNet >= " & val(TxtNetValue.text) & ""
           ElseIf RdNet(4).value = True Then
                StrWhere = StrWhere + " and QryTransactionsTotal.TransNet  <= " & val(TxtNetValue.text) & ""
            End If
                Begin = True
            End If
        End If
        
     '''''''''''''''''''
        If TxtManualNO.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and dbo.Transactions.ManualNO like '%" & (TxtManualNO.text) & "%'"
            Else
                StrWhere = StrWhere + " where dbo.Transactions.ManualNO  like '%" & (TxtManualNO.text) & "%'"
                Begin = True
            End If
        End If
        If DealingForm = InventoryOut Then
        If DCEquipments.text <> "" And val(DCEquipments.BoundText) <> 0 Then
            If Begin = True Then
                StrWhere = StrWhere + " and dbo.Transactions.FixesAssetsID =" & val(DCEquipments.BoundText) & ""
            Else
                StrWhere = StrWhere + " where dbo.Transactions.FixesAssetsID =" & val(DCEquipments.BoundText) & ""
                Begin = True
            End If
         End If
         
        End If
           If TxtCashCustomerName.text <> "" Then
               If Begin = True Then
                StrWhere = StrWhere + " and dbo.Transactions.CashCustomerName like '%" & (TxtCashCustomerName.text) & "%'"
            Else
                StrWhere = StrWhere + " where dbo.Transactions.CashCustomerName like '%" & (TxtCashCustomerName.text) & "%'"
                Begin = True
            End If
      End If
        If TxtCashCustomerPhone.text <> "" Then
               If Begin = True Then
                StrWhere = StrWhere + " and dbo.Transactions.CashCustomerPhone like '%" & (TxtCashCustomerPhone.text) & "%'"
            Else
                StrWhere = StrWhere + " where dbo.Transactions.CashCustomerPhone like '%" & (TxtCashCustomerPhone.text) & "%'"
                Begin = True
            End If
      End If
          If Not SystemOptions.IsHiddenUser Then
        StrWhere = StrWhere & " and IsNull(Transactions.IsHiddenInv,0) =0"
     
     End If
              If Trim(order_no.text) <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and transactions.order_no=N'" & (order_no.text) & "'"
            Else
                StrWhere = StrWhere + " where transactions.order_no=N'" & Trim(order_no.text) & "'"
                Begin = True
            End If
        End If
        

        If XPTxtBillNum.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and  NoteSerial1  like '%" & (XPTxtBillNum.text) & "%'"
            Else
                StrWhere = StrWhere + " where NoteSerial1  like '%" & (XPTxtBillNum.text) & "%'"
                Begin = True
            End If
        End If

        If txtOldNoteSerial1.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and  OldNoteSerial1  like '%" & (txtOldNoteSerial1.text) & "%'"
            Else
                StrWhere = StrWhere + " where OldNoteSerial1  like '%" & (txtOldNoteSerial1.text) & "%'"
                Begin = True
            End If
        End If

        If Me.CboPayMentType.ListIndex <> -1 Then
            If Me.CboPayMentType.ListIndex = 0 Then
                StrWhere = StrWhere + " and Transactions.PaymentType=0 "
            ElseIf Me.CboPayMentType.ListIndex = 1 Then
                StrWhere = StrWhere + " and Transactions.PaymentType=1"
                  ElseIf Me.CboPayMentType.ListIndex = 3 Then
                StrWhere = StrWhere + " and Transactions.PaymentType=2"
                
            End If
        End If
   If Me.DCPaymentNet.BoundText <> "" And DCPaymentNet.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.Doctype =" & Me.DCPaymentNet.BoundText & ""
            Else
                StrWhere = StrWhere + " where Transactions.Doctype =" & Me.DCPaymentNet.BoundText & ""
                Begin = True
            End If
        End If

        If DCboClientsName.BoundText <> "" And DCboClientsName.text <> "" Then
            
            
            XPChkSearchType.value = Checked

            If XPChkSearchType.value = Checked Then
                If Begin = True Then
                If Me.DealingForm = internalorder Then
                    StrWhere = StrWhere + " and Transactions.CusID1 =" & Trim(DCboClientsName.BoundText)
                ElseIf Me.DealingForm = InvoiceTransactionCompose Then
                    StrWhere = StrWhere + " and Transactions.FarmID =" & Trim(DCboClientsName.BoundText)
                Else
                StrWhere = StrWhere + " and Transactions.CusID =" & Trim(DCboClientsName.BoundText)
                End If
                
                Else
                    StrWhere = StrWhere + " where Transactions.CusID =" & Trim(DCboClientsName.BoundText)
                    Begin = True
                End If

            Else

                If Begin = True Then
                    StrWhere = StrWhere
                Else
                    StrWhere = StrWhere + " where CusName LIKE'" & Trim(DCboClientsName.text) & "%'"
                    Begin = True
                End If
            End If
            
            
        End If
        If Me.DcboUsers.BoundText <> "" And DcboUsers.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.UserID =" & Me.DcboUsers.BoundText & ""
            Else
                StrWhere = StrWhere + " where Transactions.UserID =" & Me.DcboUsers.BoundText & ""
                Begin = True
            End If
        End If

        If Not IsNull(Me.DTPFrom.value) Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.Transaction_date >=" & SQLDate(Me.DTPFrom.value, True) & ""
            Else
                StrWhere = StrWhere + " where Transactions.Transaction_date >=" & SQLDate(Me.DTPFrom.value, True) & ""
                Begin = True
            End If
        End If

        If Not IsNull(Me.DTPTo.value) Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.Transaction_date <=" & SQLDate(Me.DTPTo.value, True) & ""
            Else
                StrWhere = StrWhere + " where Transactions.Transaction_date <=" & SQLDate(Me.DTPTo.value, True) & ""
                Begin = True
            End If
        End If

        If Me.DcboStores.BoundText <> "" And DcboStores.text <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.StoreID=" & Me.DcboStores.BoundText & ""
            Else
                StrWhere = StrWhere + " where Transactions.StoreID=" & Me.DcboStores.BoundText & ""
                Begin = True
            End If
        End If
 If Trim(txtContainerNo.text) <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.ContainerNo ='" & Trim(txtContainerNo.text) & "'"
            Else
                StrWhere = StrWhere + " where Transactions.ContainerNo ='" & Trim(txtContainerNo.text) & "'"
                Begin = True
            End If
        End If
        
          If val(txtPrevValue.text) <> 0 Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.CarPrevValue =" & val(txtPrevValue.text) & ""
            Else
                StrWhere = StrWhere + " where Transactions.CarPrevValue =" & val(txtPrevValue.text) & ""
                Begin = True
            End If
        End If
        
        
       

  If Trim(TxtPlatNo.text) <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transactions.PlateNo ='" & Trim(TxtPlatNo.text) & "'"
            Else
                StrWhere = StrWhere + " where Transactions.PlateNo ='" & Trim(TxtPlatNo.text) & "'"
                Begin = True
            End If
        End If
         
        
        If Me.DCboItem.text <> "" Then

            'If Me.DCboItem.BoundText <> "" Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transaction_Details.Item_ID=" & Me.DCboItem.BoundText & ""
            Else
                StrWhere = StrWhere + " where Transaction_Details.Item_ID=" & Me.DCboItem.BoundText & ""
                Begin = True
            End If
        End If

        If val(TxtItemQty.text) > 0 Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transaction_Details.Quantity=" & val(TxtItemQty.text) & ""
            Else
                StrWhere = StrWhere + " where Transaction_Details.Quantity=" & val(TxtItemQty.text) & ""
                Begin = True
            End If
        End If

        If val(TxtItemPrice.text) > 0 Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transaction_Details.Price=" & val(TxtItemPrice.text) & ""
            Else
                StrWhere = StrWhere + " where Transaction_Details.Price=" & val(TxtItemPrice.text) & ""
                Begin = True
            End If
        End If

        If Trim(Me.TxtItemSerial.text) <> "" Then
            If ChkSerialSearchType.value = vbChecked Then
                If Begin = True Then
                    StrWhere = StrWhere + " and Transaction_Details.ItemSerial='" & Trim(TxtItemSerial.text) & "'"
                Else
                    StrWhere = StrWhere + " where Transaction_Details.ItemSerial='" & Trim(TxtItemSerial.text) & "'"
                    Begin = True
                End If

            ElseIf ChkSerialSearchType.value = vbUnchecked Then

                If Begin = True Then
                    StrWhere = StrWhere + " and Transaction_Details.ItemSerial like '%" & Trim(TxtItemSerial.text) & "%'"
                Else
                    StrWhere = StrWhere + " where Transaction_Details.ItemSerial like '%" & Trim(TxtItemSerial.text) & "%'"
                    Begin = True
                End If
            End If
        End If


   '   If SystemOptions.usertype <> UserAdminAll Then
    If m_SearchFrom = 0 And localindex = 4 Then
    Else
      StrSQL = StrSQL & "  AND      Transactions.BranchId in(" & Current_branchSql & ")"
    End If
   '             StrWhere = StrWhere & " AND   Transactions.BranchId=" & Current_branch
   '         End If
            
            

     If SystemOptions.usertype <> UserAdminAll Then
 
          If SystemOptions.FixedCustomer = 1 Then
          If m_SearchFrom = 0 And localindex = 4 Then
          Else
            StrWhere = StrWhere & " and    Transactions.UserID = " & user_id
        End If
             End If
    End If
   If (DealingForm = InvoiceTransaction Or DealingForm = PurchaseTransaction Or DealingForm = InventoryOut) And TXTOrDer_no.text <> "" Then
         If Begin = True Then
                    StrWhere = StrWhere + " and Transactions.order_no like '%" & Trim(TXTOrDer_no.text) & "%'"
                Else
                    StrWhere = StrWhere + " where Transactions.order_no like '%" & Trim(TXTOrDer_no.text) & "%'"
                    Begin = True
                End If
    End If
    
    If Index = 310 Then
        If val(mmItemId) <> 0 Then
            MySQL = MySQL + "                     and  dbo.Transaction_Details.Item_ID = " & val(mmItemId)
        End If
            If mCusId <> 0 Then
                MySQL = MySQL + "                     and  dbo.Transactions.CusId = " & val(mCusId)
            End If
    End If
    If TxtTransactionComment.text <> "" Then
        StrWhere = StrWhere + " and Transactions.TransactionComment like '%" & Trim(TxtTransactionComment.text) & "%'"
    
    End If
        Build_Sql = StrSQL + StrWhere + " order by Transactions.NoteSerial1 "
    End If

    Exit Function
ErrTrap:
End Function

Public Property Get DealingForm() As GridTransType
    DealingForm = m_DealingForm
End Property

Public Property Let DealingForm(ByVal vNewValue As GridTransType)
    'If vNewValue = OpeningBalance Or vNewValue = PurchaseTransaction Or vNewValue = InvoiceTransaction Then
    m_DealingForm = vNewValue
    'End If
End Property

Public Property Get ExtraRetrunObject() As Object
    Set ExtraRetrunObject = M_ExtraRetrunObject
End Property

Public Property Set ExtraRetrunObject(ByVal vNewValue As Object)
    'ﬁ„  »⁄„· Â–Â «·Œ«’Ì… „Œ’Ê’ Õ Ï Ì„ﬂ‰‰Ï
    '«‰ «” Œœ„ ‘«‘… «·»ÕÀ ⁄‰ «·Õ—ﬂ«  «· Ã«—Ì…
    '„‰ Œ·«· ‘«‘… «·„ﬁ»Ê÷«  ÕÌÀ Ì„ﬂ‰‰Ï
    '«‰ «” —Ã⁄ ﬂÊœ «·Õ—ﬂ… «· Ã«—Ì…
    '›Ï ‘«‘… „À· ‘«‘… «·„ﬁ»Ê÷« 
    Set M_ExtraRetrunObject = vNewValue
End Property

Public Property Get ExtraRetrunObject1() As Object
    Set ExtraRetrunObject1 = M_ExtraRetrunObject1
End Property

Public Property Set ExtraRetrunObject1(ByVal vNewValue As Object)
    'ﬁ„  »⁄„· Â–Â «·Œ«’Ì… „Œ’Ê’ Õ Ï Ì„ﬂ‰‰Ï
    '«‰ «” Œœ„ ‘«‘… «·»ÕÀ ⁄‰ «·Õ—ﬂ«  «· Ã«—Ì…
    '„‰ Œ·«· ‘«‘… «·„ﬁ»Ê÷«  ÕÌÀ Ì„ﬂ‰‰Ï
    '«‰ «” —Ã⁄ ﬂÊœ «·Õ—ﬂ… «· Ã«—Ì…
    '›Ï ‘«‘… „À· ‘«‘… «·„ﬁ»Ê÷« 
    Set M_ExtraRetrunObject1 = vNewValue
End Property

Public Property Get ExtraRetrunObject2() As Object
    Set ExtraRetrunObject2 = M_ExtraRetrunObject2
End Property

Public Property Set ExtraRetrunObject2(ByVal vNewValue As Object)
    'ﬁ„  »⁄„· Â–Â «·Œ«’Ì… „Œ’Ê’ Õ Ï Ì„ﬂ‰‰Ï
    '«‰ «” Œœ„ ‘«‘… «·»ÕÀ ⁄‰ «·Õ—ﬂ«  «· Ã«—Ì…
    '„‰ Œ·«· ‘«‘… «·„ﬁ»Ê÷«  ÕÌÀ Ì„ﬂ‰‰Ï
    '«‰ «” —Ã⁄ ﬂÊœ «·Õ—ﬂ… «· Ã«—Ì…
    '›Ï ‘«‘… „À· ‘«‘… «·„ﬁ»Ê÷« 
    Set M_ExtraRetrunObject2 = vNewValue
End Property





Private Sub TxtItemCode_KeyDown(KeyCode As Integer, _
                                Shift As Integer)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset

    If KeyCode = vbKeyReturn Then
        If Trim(Me.TxtItemCode.text) <> "" Then
            If Trim(Me.TxtItemCode.text) = "" Then
                Exit Sub
            End If
            StrSQL = "Select ItemID From TblItems Where ItemCode='" & Trim(Me.TxtItemCode.text) & "'"
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                DCboItem.BoundText = rs("ItemID").value
            Else
                'Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
                'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
        End If
    End If

End Sub

Private Sub TxtItemPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtItemPrice.text, 0)
End Sub

Private Sub TxtItemQty_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtItemQty.text, 0)
End Sub


Private Sub TxtNetValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtNetValue.text, 0)
End Sub

Private Sub TxtTotalValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTotalValue.text, 0)
End Sub

