VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDiscounts 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  «·«‘⁄«—«   "
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   DrawWidth       =   10
   Icon            =   "FrmDiscounts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImportExcel 
      Caption         =   "«ﬂ”Ì·"
      Height          =   315
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   96
      Top             =   2670
      Width           =   1305
   End
   Begin VB.TextBox txtFiterWaiverNoteSerial 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   780
      RightToLeft     =   -1  'True
      TabIndex        =   95
      Top             =   930
      Width           =   1350
   End
   Begin VB.TextBox txtFiterWaiver 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   630
      RightToLeft     =   -1  'True
      TabIndex        =   93
      Top             =   960
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox txtORDER_NO 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   6360
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   91
      Top             =   3060
      Width           =   2430
   End
   Begin VB.CheckBox chkTaxExempt 
      Caption         =   "„⁄›«…"
      Height          =   345
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   90
      Top             =   1740
      Width           =   690
   End
   Begin VB.TextBox TXTIban 
      Height          =   495
      Left            =   -360
      TabIndex        =   87
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox txt_Currency_rate 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   84
      Text            =   "1"
      Top             =   2565
      Width           =   765
   End
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   30
      ScaleHeight     =   3675
      ScaleWidth      =   3405
      TabIndex        =   80
      Top             =   6780
      Width           =   3465
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H0080FFFF&
      Caption         =   "»Ì«‰«  ›Ê« Ì— «·„»Ì⁄« "
      Height          =   6135
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   72
      Top             =   3990
      Visible         =   0   'False
      Width           =   10455
      Begin VB.CommandButton Command10 
         BackColor       =   &H8000000B&
         Caption         =   "«·€«¡ «·”œ«œ"
         Height          =   315
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   240
         Width           =   1695
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   4860
         Left            =   0
         TabIndex        =   74
         Top             =   600
         Width           =   10320
         _cx             =   18203
         _cy             =   8572
         Appearance      =   2
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmDiscounts.frx":038A
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
         ExplorerBar     =   1
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
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Height          =   375
         Left            =   9840
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   -240
         RightToLeft     =   -1  'True
         TabIndex        =   76
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   5640
         Width           =   8295
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «·›Ê« Ì—"
         Height          =   255
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   75
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   5640
         Width           =   1575
      End
   End
   Begin VB.TextBox TxtVATYou 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   840
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   2160
      Width           =   1050
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5370
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   2130
      Width           =   1050
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   7680
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   2160
      Width           =   1050
   End
   Begin VB.TextBox TxtVAT 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   3000
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   2160
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   7680
      TabIndex        =   62
      Top             =   1800
      Width           =   1050
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox TxtSearchCode 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   7680
      TabIndex        =   59
      Top             =   1440
      Width           =   1050
   End
   Begin VB.TextBox TxtTransID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   840
      Visible         =   0   'False
      Width           =   555
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
      Height          =   885
      Index           =   1
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   4650
      Width           =   10335
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   240
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   43
         Top             =   180
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboCreditSide 
         Height          =   315
         Left            =   90
         TabIndex        =   44
         Top             =   510
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   255
         Index           =   7
         Left            =   6120
         TabIndex        =   53
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   450
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   11
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   870
         Width           =   1485
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3630
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   450
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·› —… :"
         Height          =   315
         Index           =   29
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   315
         Index           =   30
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› œ«∆‰"
         Height          =   285
         Index           =   31
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   510
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› „œÌ‰"
         Height          =   285
         Index           =   32
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.CheckBox ChkTrans 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰ Õ”«» ›« Ê—…"
      Height          =   225
      Left            =   1950
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox CboDiscountType 
      Height          =   315
      Left            =   7320
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   975
      Left            =   3000
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3570
      Width           =   5835
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   360
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   930
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.ComboBox DCboCashType 
      Height          =   315
      Left            =   4800
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   1065
      Index           =   0
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox TxtTransSerial 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   1005
      End
      Begin VB.ComboBox CboTrans 
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   1995
      End
      Begin ImpulseButton.ISButton CmdSearchTrans 
         Height          =   345
         Left            =   600
         TabIndex        =   7
         Top             =   630
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
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
         ButtonImage     =   "FrmDiscounts.frx":0661
      End
      Begin ImpulseButton.ISButton CmdOpenTrans 
         Height          =   345
         Left            =   90
         TabIndex        =   10
         Top             =   630
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
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
         ButtonImage     =   "FrmDiscounts.frx":09FB
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«œŒ· —ﬁ„ «·›« Ê—…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   10
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   690
         Width           =   1305
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Œ — ‰Ê⁄ «·›« Ê—…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   12
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   1305
      End
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   585
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   10335
      _cx             =   18230
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   18
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
      Caption         =   "  «·«‘⁄«—«   "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2970
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox DefaultInvoicetype 
         Height          =   315
         ItemData        =   "FrmDiscounts.frx":0D95
         Left            =   6450
         List            =   "FrmDiscounts.frx":0D97
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   30
         Width           =   1890
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1185
         TabIndex        =   16
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
         ButtonImage     =   "FrmDiscounts.frx":0D99
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
         Left            =   120
         TabIndex        =   17
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
         ButtonImage     =   "FrmDiscounts.frx":1133
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
         Left            =   1710
         TabIndex        =   18
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
         ButtonImage     =   "FrmDiscounts.frx":14CD
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
         Left            =   645
         TabIndex        =   19
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
         ButtonImage     =   "FrmDiscounts.frx":1867
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin MSDataListLib.DataCombo DCDocTypes 
         Height          =   315
         Left            =   3510
         TabIndex        =   82
         Top             =   90
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
         Left            =   4680
         TabIndex        =   83
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   2520
         Picture         =   "FrmDiscounts.frx":1C01
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   4770
      TabIndex        =   0
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   248381441
      CurrentDate     =   41640
   End
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   6240
      TabIndex        =   20
      Top             =   5580
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   1560
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6000
      Width           =   8745
      _cx             =   15425
      _cy             =   953
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   7875
         TabIndex        =   22
         Top             =   105
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   1
         Left            =   6870
         TabIndex        =   23
         Top             =   105
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   2
         Left            =   5955
         TabIndex        =   24
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   3
         Left            =   5055
         TabIndex        =   25
         Top             =   105
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   4
         Left            =   4155
         TabIndex        =   26
         Top             =   105
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   27
         Top             =   105
         Width           =   870
         _ExtentX        =   1535
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   375
         Left            =   1170
         TabIndex        =   28
         Top             =   105
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   5
         Left            =   3255
         TabIndex        =   29
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   8
         Left            =   2280
         TabIndex        =   63
         Top             =   105
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…"
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
   End
   Begin MSDataListLib.DataCombo Dcbranch 
      Bindings        =   "FrmDiscounts.frx":5869
      Height          =   315
      Left            =   0
      TabIndex        =   55
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
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
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   120
      TabIndex        =   57
      Top             =   6120
      Width           =   1155
      _ExtentX        =   2037
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
   Begin ImpulseButton.ISButton CmdCusSearch 
      Height          =   345
      Left            =   2400
      TabIndex        =   58
      Top             =   1440
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
      ButtonImage     =   "FrmDiscounts.frx":587E
      DrawFocusRectangle=   0   'False
   End
   Begin MSDataListLib.DataCombo DcEmp 
      Bindings        =   "FrmDiscounts.frx":5C18
      Height          =   315
      Left            =   3000
      TabIndex        =   67
      Top             =   1800
      Width           =   4635
      _ExtentX        =   8176
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
   Begin VB.TextBox TxtValueTemp 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   79
      Top             =   1440
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "⁄—÷ «·›Ê« Ì—"
      Height          =   315
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   78
      Top             =   960
      Width           =   1695
   End
   Begin MSDataListLib.DataCombo DcCurrency 
      Height          =   315
      Left            =   1680
      TabIndex        =   85
      Top             =   2550
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
      TabIndex        =   89
      Top             =   2880
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   635
      _Version        =   393216
      Format          =   242941953
      CurrentDate     =   38784
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ’›Ì…"
      Height          =   255
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   94
      Top             =   1020
      Width           =   435
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·›« Ê—…"
      Height          =   255
      Left            =   8940
      RightToLeft     =   -1  'True
      TabIndex        =   92
      Top             =   3060
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Õ”«» «·»‰ﬂÌ"
      Height          =   480
      Index           =   136
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   88
      Top             =   0
      Width           =   840
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "««·⁄„·…"
      Height          =   300
      Index           =   65
      Left            =   2010
      RightToLeft     =   -1  'True
      TabIndex        =   86
      Top             =   2580
      Width           =   1080
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰”»… «·VAT"
      Height          =   285
      Index           =   14
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁÌ„… «·‘«„·…"
      Height          =   255
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„…  «·VAT"
      Height          =   285
      Index           =   13
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„‰œÊ»"
      Height          =   255
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Lb_note_value_by_characters 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   2640
      Width           =   8775
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   285
      Index           =   5
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   3690
      Width           =   1485
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·”‰œ"
      Height          =   285
      Index           =   4
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì· √Ê «·„Ê—œ"
      Height          =   285
      Index           =   3
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   1440
      Width           =   1485
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌ„…"
      Height          =   285
      Index           =   2
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   5700
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   615
      Width           =   1485
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄„Ì· «Ê „Ê—œ"
      Height          =   285
      Index           =   0
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   960
      Width           =   1485
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   9345
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   5535
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   5640
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   810
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   5640
      Width           =   1065
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   5640
      Width           =   825
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·«‘⁄«—"
      Height          =   285
      Index           =   9
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   960
      Width           =   1365
   End
End
Attribute VB_Name = "FrmDiscounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim FlgBillBuy As Boolean
Public mIsNoMsg As Boolean
Dim s As String
Dim zatcaStatus As Integer
Dim mIndexVat As Integer
Dim Export As Integer
Public mTypeInvoice As Integer

Public FlgAproved As Integer
Private Const AR_ACCT_SERIAL As Long = 8427   ' ”Ì—Ì«· Õ”«» «·–„„ „‰ ÃœÊ· ACCOUNTS

 Function saveBillBuy()
    Dim StrSQL As String
   ' Dim StrSQL  As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Diff = 0
Dim RsDetails As ADODB.Recordset
      If Me.TxtModFlg.text = "E" Then
    StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.XPTxtID.text) & " and TransType=2"
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.text) & " and TransType=2"
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblNotesBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
    TxtValueTemp.text = val(TxtTotal.text)
    For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID1").value = val(XPTxtID.text)
            RsDetails("TransType").value = 2
            RsDetails("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
            RsDetails("branch_no").value = val(.TextMatrix(i, .ColIndex("branch_no")))
            RsDetails("NoteSerial1").value = val(.TextMatrix(i, .ColIndex("NoteSerial1")))
            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
            Note_Value1 = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            Diff = 0
            If val(TxtValueTemp.text) > 0 Then
          If val(TxtValueTemp.text) <= Note_Value1 Then
          Diff = val(TxtValueTemp.text)
          TxtValueTemp.text = val(TxtValueTemp.text) - Note_Value1
          Else
          Diff = Note_Value1
          TxtValueTemp.text = val(TxtValueTemp.text) - Note_Value1
          End If
            End If
          ' .TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("RemainingValue")))
            .TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
            
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
            
            RsDetails("too").value = (.TextMatrix(i, .ColIndex("too")))
            RsDetails("NoteDate").value = IIf((.TextMatrix(i, .ColIndex("NoteDate"))) = "", Null, (.TextMatrix(i, .ColIndex("NoteDate"))))
            If .TextMatrix(i, .ColIndex("DueDate")) <> "" And .TextMatrix(i, .ColIndex("DueDate")) <> " " Then
            RsDetails("DueDate").value = IIf((.TextMatrix(i, .ColIndex("DueDate"))) = "", Null, (.TextMatrix(i, .ColIndex("DueDate"))))
            Else
            RsDetails("DueDate").value = Null
            End If
            RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
            RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            RsDetails.update
                
            If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
            StrSQL = "Update Transactions Set  TotalPayed=1 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
             Else
                 StrSQL = "Update Transactions Set  TotalPayed=0 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
      End If
    Next i
End With
    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
    For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.text)
            RsDetails("RecDate").value = XPDtbTrans.value
            RsDetails("Serial").value = TxtNoteSerial1.text
            RsDetails("TransType").value = 2
            RsDetails("Transaction_ID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails.update
        End If
    Next i
End With

End Function
Private Sub CboDiscountType_Change()
lbl(2).Visible = True
lbl(13).Visible = True
lbl(14).Visible = True
XPTxtVal.Visible = True
TxtVAT.Visible = True
TxtVATYou.Visible = True
Label3.Visible = True
lbl(3).Visible = True
TxtSearchCode.Visible = True
DBCboClientName.Visible = True
CmdCusSearch.Visible = True
Text1.Visible = True
DcEmp.Visible = True
If val(CboDiscountType.ListIndex) = 0 And SystemOptions.AllowDiscountAllowedFIFO = True Then
Command1.Visible = True
BillCustomer
Else
Command1.Visible = False
End If
If val(CboDiscountType.ListIndex) = 3 Then
    mIndexVat = 17
Else
    mIndexVat = 23
End If
If val(CboDiscountType.ListIndex) = 5 Or val(CboDiscountType.ListIndex) = 6 Then

If SystemOptions.UserInterface = ArabicInterface Then
Label1.Caption = "«·ﬁÌ„… "
End If
lbl(2).Visible = False
lbl(13).Visible = False
lbl(14).Visible = False
XPTxtVal.Visible = False
TxtVAT.Visible = False
TxtVATYou.Visible = False
Label3.Visible = False
'lbl(3).Visible = False
'TxtSearchCode.Visible = False
'DBCboClientName.Visible = False
'CmdCusSearch.Visible = False
Text1.Visible = False
DcEmp.Visible = False
Else
If SystemOptions.UserInterface = ArabicInterface Then
Label1.Caption = "«·ﬁÌ„… «·‘«„·…"
End If
End If
If val(CboDiscountType.ListIndex) = 3 Or val(CboDiscountType.ListIndex) = 4 Then
DcboDebitSide.Enabled = True
DcboCreditSide.Enabled = True
DCboCashType.Visible = True
TxtSearchCode.Visible = False
Text1.Visible = False
'DBCboClientName.Visible = False
DcEmp.Visible = False
CmdCusSearch.Visible = False
Label3.Visible = False
lbl(3).Visible = False
ElseIf val(CboDiscountType.ListIndex) = 5 Then
DcboCreditSide.Enabled = False
DcboDebitSide.Enabled = True
ElseIf val(CboDiscountType.ListIndex) = 6 Then
DcboCreditSide.Enabled = True
DcboDebitSide.Enabled = False
Else
lbl(3).Visible = True
Label3.Visible = True
CmdCusSearch.Visible = True
DcEmp.Visible = True
DBCboClientName.Visible = True
Text1.Visible = True
TxtSearchCode.Visible = True
DCboCashType.Visible = True
DcboDebitSide.Enabled = False
DcboCreditSide.Enabled = False
End If
    WriteDev
    Calculte
    
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
      
        SaveQRCode "Notes", "NoteID", val(XPTxtID), TxtNoteSerial1.text, (XPDtbTrans.value), _
        (TxtTotal.text), Picture1, 0, (TxtVAT.text), (TxtTotal.text)
        
 
        
    
MySQL = " SELECT    N'" & Trim(DcboDebitSide.text) & "' as DcboDebitSide,TblCustemers.VATNO, N'" & Trim(DcboCreditSide.text) & "' as DcboCreditSide,  dbo.Notes.NoteID, dbo.Notes.NoteType, dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Notes.NoteSerial,"
MySQL = MySQL & "                       Notes.QrCodeImage, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value, dbo.Notes.NoteHijriDate, dbo.Notes.note_value_by_characters, dbo.Notes.Remark, dbo.Notes.NoteDate,"
MySQL = MySQL & "                      dbo.Notes.Member_ID, dbo.Notes.Transaction_ID, dbo.Notes.BankID, dbo.Notes.CashingType, dbo.Notes.numbering_type, dbo.Notes.sanad_year,"
MySQL = MySQL & "                      dbo.Notes.sanad_month, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Cus_Phone,"
MySQL = MySQL & "                      dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.E_mail, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.Fullcode,"
MySQL = MySQL & "                      dbo.TblCustemers.CustGID, dbo.TblCustemers.Mobile2,TblCustemers.VATNO, dbo.TblCustemers.Mobile1, dbo.TblCustemers.HomeTel, dbo.TblCustemers.JobTelConvert,"
MySQL = MySQL & "                      dbo.TblCustemers.JobTel, dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTitle, dbo.TblCustemers.Company, dbo.Notes.EmpId, dbo.TblEmployee.Emp_Name,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality,"
MySQL = MySQL & "                      dbo.TblEmployee.Fullcode AS Expr1, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee3 , dbo.TblEmployee.Emp_Namee4 ,Notes.ORDER_NO, dbo.Notes.VAT, dbo.Notes.VATYou, dbo.Notes.TotalValue"
MySQL = MySQL & " FROM         dbo.Notes LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.Notes.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " WHERE     ((dbo.Notes.NoteType = 9) or (dbo.Notes.NoteType = 9089) or (dbo.Notes.NoteType = 9090) or (dbo.Notes.NoteType = 9099) OR (dbo.Notes.NoteType = 9082) or (dbo.Notes.NoteType = 9083)or "
MySQL = MySQL & "                      (dbo.Notes.NoteType = 10) OR"
MySQL = MySQL & "                      (dbo.Notes.NoteType = 8034))AND (dbo.Notes.NoteID = " & val(XPTxtID.text) & ")"
If val(CboDiscountType.ListIndex) = 3 Or val(CboDiscountType.ListIndex) = 4 Then
 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllowDiscount2.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllowDiscount2E.rpt"
       End If
 Else
  If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllowDiscount.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllowDiscountE.rpt"
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
Dim cOptions        As ClsCompanyInfo
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
        
        StrReportTitle = ""
    End If

    
    If SystemOptions.VATNoAccordActivity = False Then
        xReport.ParameterFields(5).AddCurrentValue cCompanyInfo.VATRegNo
    Else
        xReport.ParameterFields(5).AddCurrentValue GetRegVATNo(val(Dcbranch.BoundText))
    End If
    
    
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".")
    
    xReport.ParameterFields(6).AddCurrentValue WriteNo(Format(val(Me.TxtVATYou.text) + val(Me.TxtTotal.text), "0.00"), 0, True, ".")
    xReport.ParameterFields(7).AddCurrentValue WriteNo(Format(Me.TxtTotal.text, "0.00"), 0, True, ".")
'    xReport.ParameterFields(6).AddCurrentValue WriteNo(Format(Me.XPTxtCurrent.text, "0.00"), 0, True, ".")
'    xReport.ParameterFields(7).AddCurrentValue WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".")
'    xReport.ParameterFields(8).AddCurrentValue WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".")


'xReport.ParameterFields(8).AddCurrentValue Lb_note_value_by_characters
    
    
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
Private Sub CboDiscountType_Click()
    CboDiscountType_Change
End Sub

Private Sub ChkTrans_Click()
    Me.lbl(10).Enabled = ChkTrans.value
    Me.lbl(12).Enabled = ChkTrans.value
    Me.CboTrans.Enabled = ChkTrans.value
    Me.TxtTransID.Enabled = ChkTrans.value
    Me.TxtTransSerial.Enabled = ChkTrans.value
    Me.CmdSearchTrans.Enabled = ChkTrans.value
    Me.CmdOpenTrans.Enabled = ChkTrans.value

End Sub

Public Sub Cmd_Click(Index As Integer)
     On Error GoTo ErrTrap
    Dim Msg As String
        If CboDiscountType.ListIndex = 3 Or CboDiscountType.ListIndex = 4 Then
            If (Index = 1 Or Index = 4) And zatcaStatus = 1 Then
                    Msg = "·« Ì„ﬂ‰  ⁄œÌ· «Ê Õ–› «Ì „” ‰œ Ì„ﬂ‰ﬂ ⁄„· „” ‰œ ⁄ﬂ”Ì ›ﬁÿ"
                        Msg = Msg & CHR(13) & ""
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
            End If
        End If
       If mZakamsg <> "" Then
            
        MsgBox mZakamsg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, " ÂÌ∆… «·“ﬂ«… Ê«·÷—Ì»… Ê«·Ã„«—ﬂ «·„—Õ·… «·À«‰Ì… - „—Õ·… «·—»ÿ Ê«· ﬂ«„·"
    End If
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
          
            '   Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=9 OR NoteType=10"))
            Me.DCboUserName.BoundText = user_id
            XPDtbTrans.SetFocus
            Me.Dcbranch.BoundText = branch_id
            Calculte
            DefaultInvoicetype.ListIndex = SystemOptions.DefaultInvoicetype
            zatcaStatus = 0
            FlgAproved = 0
            txtDateRec.value = Date
        Case 1
         If CboDiscountType.ListIndex = 3 Or CboDiscountType.ListIndex = 4 Then
            If checkCustomerdata(val(Me.DBCboClientName.BoundText), val(XPTxtVal), val(DefaultInvoicetype.ListIndex), DcCurrency.text, Export) = False Then Exit Sub
        End If
    If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
       If SystemOptions.AllowDiscountAllowedFIFO = True And val(DBCboClientName.BoundText) <> 0 And val(CboDiscountType.ListIndex) = 0 Then
          Command10_Click
          BillCustomer
       End If
            Me.DCboUserName.BoundText = user_id
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
              
              
            

            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·«  "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Dcbranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
    
            my_branch = Me.Dcbranch.BoundText

            '       If Me.TxtModFlg.text = "N" Then
 
            '   End If
  If val(TxtVAT.text) > 0 And (CboDiscountType.ListIndex = 3 Or CboDiscountType.ListIndex = 4) Then
If GetValueAddedAccount(XPDtbTrans.value, , , 1, mIndexVat) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì…"
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If
If val(CboDiscountType.ListIndex) = 0 And SystemOptions.AllowDiscountAllowedFIFO = True Then
BillCustomer
AutoCalculate
End If
         
            SaveData

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

            Load FrmDiscountsSearch
             FrmDiscountsSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200
        Case 8
        print_report

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAttach_Click()
     On Error Resume Next
ShowAttachments TxtNoteSerial1, "0712201406", XPTxtID

End Sub

Private Sub CmdCusSearch_Click()
  Load FrmCustemerSearch
            FrmCustemerSearch.SearchType = -2
            FrmCustemerSearch.RetrunType = 1
            Set FrmCustemerSearch.DcboCustomers = Me.DBCboClientName
            FrmCustemerSearch.show vbModal

End Sub

Private Sub CmdSearchTrans_Click()
    Dim Msg As String

    If Me.CboTrans.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·Õ—ﬂ… «·„—«œ «·»ÕÀ ⁄‰Â«..."
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboTrans.SetFocus
        Sendkeys "{F4}"
        Exit Sub
    End If

    If Me.CboTrans.ListIndex = 0 Then
        '›« Ê—… „»Ì⁄« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = InvoiceTransaction
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPayMentType.ListIndex = 1
        FrmBuySearch.CboPayMentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show
    ElseIf Me.CboTrans.ListIndex = 1 Then
        '›« Ê—… „‘ —Ì« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = PurchaseTransaction
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPayMentType.ListIndex = 1
        FrmBuySearch.CboPayMentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… ‘—«¡"
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show vbModal
    ElseIf Me.CboTrans.ListIndex = 2 Then
        '›« Ê—… „— Ã⁄ „»Ì⁄« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = ReturnSalling
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPayMentType.ListIndex = 1
        FrmBuySearch.CboPayMentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… „— Ã⁄ „»Ì⁄« "
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show
    ElseIf Me.CboTrans.ListIndex = 3 Then
        '›« Ê—… „— Ã⁄ „‘ —Ì« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = Returntransaction
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPayMentType.ListIndex = 1
        FrmBuySearch.CboPayMentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ „— Ã⁄ «·„‘ —Ì« "
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show vbModal

    ElseIf Me.CboTrans.ListIndex = 4 Then
        '›« Ê—… ’Ì«‰…
        Load FrmMaintanenceSearch
        Set FrmMaintanenceSearch.ExtraRetrunObject = Me.TxtTransID
        FrmMaintanenceSearch.CboPayMentType.ListIndex = 1
        FrmMaintanenceSearch.CboPayMentType.Enabled = False
        FrmMaintanenceSearch.show vbModal
    End If

End Sub

Private Sub Command1_Click()
If val(DBCboClientName.BoundText) <> 0 Then
Frame12.Visible = True
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄„Ì· «Ê·«"
Else
MsgBox "Please Select Customer"
End If
DBCboClientName.SetFocus
Exit Sub
End If
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Dim mType As Long
        If DCboCashType.ListIndex = 0 Then
            mType = 111
        ElseIf DCboCashType.ListIndex = 1 Then
            mType = 2222
        ElseIf DCboCashType.ListIndex = 2 Then
            mType = 333
        End If
        FrmCustemerSearch.SearchType = mType
        FrmCustemerSearch.show vbModal

    End If
 


End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub


Private Sub Command10_Click()
Dim i As Integer
Dim StrSQL As String
If Me.TxtModFlg.text = "E" Then
DeleteBillBuy
VSFlexGrid1.Enabled = True
        
      StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.XPTxtID.text) & " and TransType=2"
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.text) & " and TransType=2"
    Cn.Execute StrSQL, , adExecuteNoRecords

            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
VSFlexGrid1.rows = 1

FlgBillBuy = True
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·€«¡ «·”œ«œ"
Else
MsgBox "Done"
End If
    With Me.VSFlexGrid1

            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i


        End With
End If
End Sub

Private Sub DBCboClientName_Change()
    Dim fullcode As String
    Dim DefaultSalesPersonId As Integer
 
If DCboCashType.ListIndex = 0 Then
GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode, 1
ElseIf DCboCashType.ListIndex = 1 Then
GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode, 2
ElseIf DCboCashType.ListIndex = 1 Then
GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode, 3
End If

    TxtSearchCode.text = fullcode

     
            If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            
            DcEmp.BoundText = DefaultSalesPersonId
            End If
    WriteDev
If val(CboDiscountType.ListIndex) = 0 And SystemOptions.AllowDiscountAllowedFIFO = True Then
BillCustomer
End If
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    WriteDev
End Sub

Private Sub DCboCashType_Change()
    On Error Resume Next
    Dim Dcombos As New ClsDataCombos

    Select Case DCboCashType.ListIndex

        Case 0
            Dcombos.GetCustomersSuppliers 156, Me.DBCboClientName, False

        Case 1
            Dcombos.GetCustomersSuppliers 257, Me.DBCboClientName, False
        Case 2
            Dcombos.GetCustomersSuppliers 33, Me.DBCboClientName, False
    End Select

    cSearchDcbo.Refresh
End Sub

Private Sub DCboCashType_Click()
    DCboCashType_Change
End Sub

Private Sub DcboCreditSide_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 204

    End If
End Sub

Private Sub DcboDebitSide_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 203

    End If
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
     TxtNoteSerial1.text = ""
End Sub


 
 


Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ  «·”‰œ  " & TxtTransID.text & CHR(13) & "   «·›—⁄ " & Dcbranch & CHR(13) & "   «· «—ÌŒ " & XPDtbTrans & CHR(13) & "   ‰Ê⁄ «·Œ’„ " & CboDiscountType & CHR(13) & "   ⁄„Ì· / „Ê—œ  " & DCboCashType & CHR(13) & "   «”„ «·⁄„Ì· / «·„Ê—œ  " & DBCboClientName & CHR(13) & "   ﬁÌ„… «·Œ’„  " & XPTxtVal & CHR(13) & " „·«ÕŸ«  " & XPMTxtRemarks
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Vchr No.  " & TxtTransID.text & CHR(13) & "   branch " & Dcbranch & CHR(13) & "   Date " & XPDtbTrans & CHR(13) & " Discount " & CboDiscountType & CHR(13) & "   Customer/vendor  " & DCboCashType & CHR(13) & "   «Customer/vendor Name  " & DBCboClientName & CHR(13) & " Discount Value  " & XPTxtVal & CHR(13) & " Remarks " & XPMTxtRemarks
       Dim NoteType As Integer

    If Me.CboDiscountType.ListIndex = 0 Then
        'Œ’„ „”„ÊÕ »Â
        NoteType = 9
    ElseIf Me.CboDiscountType.ListIndex = 1 Then
        'Œ’„ „ﬂ ”»
        NoteType = 10
             ElseIf Me.CboDiscountType.ListIndex = 3 Then
        'Œ’„ „ﬂ ”»
        NoteType = 9082
             ElseIf Me.CboDiscountType.ListIndex = 4 Then
        'Œ’„ „ﬂ ”»
        NoteType = 9083
        ElseIf Me.CboDiscountType.ListIndex = 5 Then
        'Œ’„ „ﬂ ”»
        NoteType = 9089
            ElseIf Me.CboDiscountType.ListIndex = 6 Then
        'Œ’„ „ﬂ ”»
        NoteType = 9090
            ElseIf Me.CboDiscountType.ListIndex = 7 Then
        'Œ’„ „ﬂ ”»
        NoteType = 9099
    End If
  
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), NoteType, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TxtNoteSerial, TxtTransID
    Else
        AddToLogFile CInt(user_id), NoteType, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtNoteSerial, TxtTransID
    End If
    
End Function

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

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
            'Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    'CmdConvert.Caption = "Convert to bill"
    Dim XPic As IPictureDisp
    CmdAttach.Caption = "Attachments"


    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Cmd(8).Caption = "Print"
    Cmd(7).Caption = "Print JL "
    Label1.Caption = "Total"
    Me.Caption = " Discounts "
    Ele.Caption = Me.Caption
    Label2.Caption = "Branch"
    lbl(4).Caption = "ID"
    lbl(1).Caption = "Date"
    lbl(9).Caption = "Type"
    lbl(0).Caption = "Cust./Vendor"
    lbl(13).Caption = "VAT"
    lbl(3).Caption = "Name"
    Label3.Caption = "Employee"
    lbl(2).Caption = "Value"
    lbl(14).Caption = "VAT %"
    lbl(5).Caption = " Remarks "

    lbl(8).Caption = " By:"
    lbl(6).Caption = "Curr Rec. "
    lbl(7).Caption = "Rec. Count:"

    Fra(1).Caption = "GL"
    lbl(30).Caption = "GL#"
    lbl(29).Caption = "Interval"
    lbl(32).Caption = "Depit"
    lbl(31).Caption = "Credit"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    'Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos
    Dim Msg As String
    
    ScreenNameArabic = "«·Œ’Ê„«  «·„”„ÊÕ… Ê«·„ﬂ ”»…"
    ScreenNameEnglish = "Allowed and Earned Discounts"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    'AddTip

        With Me.DefaultInvoicetype
            .Clear
            
             


            .AddItem " ›« Ê—… ÷—Ì»Ì…  "
            .ItemData(0) = 0
     
            .AddItem " ›« Ê—… ÷—Ì»Ì… „»”ÿ… "
            .ItemData(1) = 2
         
        End With
 StrSQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, StrSQL

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, False
    Dcombos.GetUsers Me.DCboUserName
      Dcombos.GetSalesRepData Me.DcEmp
Dcombos.GetDocTypebyid Me.DCDocTypes, 21, val(Me.Dcbranch.BoundText)

    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DBCboClientName
    Dcombos.GetAccountingCodes Me.DcboDebitSide, True
    Dcombos.GetAccountingCodes Me.DcboCreditSide, True
    Dcombos.GetBranches Me.Dcbranch

    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "Select * From Notes Where (NoteType=9 OR NoteType=10 or NoteType=8034 or NoteType=9082 or NoteType=9083 or NoteType=9089 or NoteType=9090 or NoteType=9099)"
StrSQL = StrSQL & "   AND   branch_no in(" & Current_branchSql & ")"

        If SystemOptions.usertype <> UserAdmin Then
'        StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
    
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    SetDtpickerDate Me.XPDtbTrans

    With Me.CboTrans
        .Clear
        .AddItem "›« Ê—… „»Ì⁄« "
        .AddItem "›« Ê—… „‘ —Ì« "
        .AddItem "„— Ã⁄ „»Ì⁄« "
        .AddItem "„— Ã⁄ „‘ —Ì« "
        .AddItem "’Ì«‰…"
        .AddItem "Œœ„« "
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        DCboCashType.Clear
        DCboCashType.AddItem "From Customer"
        DCboCashType.AddItem "From Supp"
         DCboCashType.AddItem "Contractor "
        
        
        With Me.CboDiscountType
            .Clear
            .AddItem "Discount allowed"
            .ItemData(0) = 1
            .AddItem "Unearned discount"
            .ItemData(1) = 2
            .AddItem "Bad debts"
            .ItemData(2) = 3
            .AddItem "Debit"
            .ItemData(3) = 4
            .AddItem "Credit"
            .ItemData(4) = 5
            .AddItem "VAT-Add"
            .ItemData(5) = 6
             .AddItem "VAT-Disciunt"
            .ItemData(6) = 7
            
             .AddItem "Property Clearance"
            .ItemData(7) = 8
            
        End With
     
    Else

        DCboCashType.Clear
        DCboCashType.AddItem "  ⁄„Ì·/„” √Ã—"
        DCboCashType.AddItem "  „Ê—œ/„«·ﬂ"
        DCboCashType.AddItem "  „ﬁ«Ê·"
        
        With Me.CboDiscountType
            .Clear
            .AddItem "Œ’„ „”„ÊÕ »Â"
            .ItemData(0) = 1
            .AddItem "Œ’„ „ﬂ ”»"
            .ItemData(1) = 2
            
             .AddItem "œÌÊ‰ „⁄œÊ„… "
            .ItemData(2) = 3
            .AddItem "«‘⁄«— „œÌ‰"
            .ItemData(3) = 4
            .AddItem "«‘⁄«— œ«∆‰"
            .ItemData(4) = 5
            .AddItem " ﬁÌ„… „÷«›…-«÷«›…"
            .ItemData(5) = 6
            .AddItem " ﬁÌ„… „÷«›…-Œ’„"
            .ItemData(6) = 7
            .AddItem " ’›Ì… «„·«ﬂ"
            .ItemData(7) = 8
        End With

    End If

    Msg = "„·ÕÊŸ…:-"
    Msg = Msg & CHR(13) & "≈–« ﬂ«‰  Â–Â «·„ﬁ»Ê÷«   Õ’Ì· ·›« Ê—… „⁄Ì‰…"
    Msg = Msg & "›ÌÃ» ⁄·Ìﬂ «‰  ﬁÊ„ » ÕœÌœ Â–Â «·›« Ê—… "
    Msg = Msg & "Õ Ï Ì „ —»ÿ ⁄„·Ì… «· Õ’Ì· Â–Â „⁄ «·›« Ê—…"
    'Me.Lbl(13).Caption = Msg

    ChkTrans.value = Unchecked
    ChkTrans_Click
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
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
        If Not mIsNoMsg Then
        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)
        Else
        IntResult = 0
        End If
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

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim StrTemp As String
    Dim DblCreditNoteValue As Double
    Dim LngDebitNoteID As Long
    Dim BeginTrans As Boolean
    Dim LngDevID As Long
    Dim RsDev As ADODB.Recordset
    Dim AccountVATCreit As String
     On Error GoTo ErrTrap
                             
                             
Dim mNoteTypeSer As Integer
            If Me.CboDiscountType.ListIndex = 0 Then
            'Œ’„ „”„ÊÕ »Â
            mNoteTypeSer = 9
        ElseIf Me.CboDiscountType.ListIndex = 1 Then
            'Œ’„ „ﬂ ”»
            mNoteTypeSer = 10
            
         ElseIf Me.CboDiscountType.ListIndex = 2 Then
            'œÌÊ‰ „⁄œÊ„…
           mNoteTypeSer = 8034
        ElseIf Me.CboDiscountType.ListIndex = 3 Then
            '„œÌ‰
            mNoteTypeSer = 9082
        ElseIf Me.CboDiscountType.ListIndex = 4 Then
          'œ«∆‰
            mNoteTypeSer = 9083
       ElseIf Me.CboDiscountType.ListIndex = 5 Then
          '
            mNoteTypeSer = 9089
       ElseIf Me.CboDiscountType.ListIndex = 6 Then
          '
            mNoteTypeSer = 9090
            
 ElseIf Me.CboDiscountType.ListIndex = 7 Then
          '
            mNoteTypeSer = 9099
            
        End If
    If Me.TxtModFlg.text <> "R" Then
        If CboDiscountType.ListIndex = -1 Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·Œ’Ê„«  "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboDiscountType.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
        If DBCboClientName.text = "" And (val(CboDiscountType.ListIndex) <> 3 And val(CboDiscountType.ListIndex) <> 4 And val(CboDiscountType.ListIndex) <> 5 And val(CboDiscountType.ListIndex) <> 6) Then
            Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· √Ê «·„Ê—œ"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DBCboClientName.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
        If val(CboDiscountType.ListIndex) = 3 Or val(CboDiscountType.ListIndex) = 4 Or val(CboDiscountType.ListIndex) = 5 Or val(CboDiscountType.ListIndex) = 6 Then
        If DcboDebitSide.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«» «·„œÌ‰"
        Else
        MsgBox "Please Select Account"
        End If
        DcboDebitSide.SetFocus
        Exit Sub
        End If
         If DcboCreditSide.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«» «·œ«∆‰"
        Else
        MsgBox "Please Select Account"
        End If
        DcboCreditSide.SetFocus
        Exit Sub
        End If
        End If
        

        If val(XPTxtVal.text) = 0 Then
            Msg = "ÌÃ» «œŒ«· «·ﬁÌ„…  "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtVal.SetFocus
            Exit Sub
        End If

        If Not IsNumeric(XPTxtVal.text) Then
            Msg = "«·ﬁÌ„…  ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtVal.SetFocus
            SelectText XPTxtVal
            Exit Sub
        End If

        If Me.ChkTrans.value = vbChecked Then
            If Me.CboTrans.ListIndex = -1 Then
                Msg = "»—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·›« Ê—…..!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                CboTrans.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Me.CboDiscountType.ListIndex = 0 Then
                If Me.CboTrans.ListIndex <> 4 And Me.CboTrans.ListIndex <> 5 And Me.CboTrans.ListIndex <> 0 And Me.CboTrans.ListIndex <> 3 Then
                    Msg = "«·Œ’„ «·„”„ÊÕ »Â ÌﬂÊ‰ „⁄ ( ›« Ê—… «·»Ì⁄ «Ê „— Ã⁄ «·„‘ —Ì«  «Ê «·’Ì«‰… «Ê «·Œœ„« )"
                    Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ ‰Ê⁄ «·Œ’„ «Ê ‰Ê⁄ «·›« Ê—… «·„Õœœ…..!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    CboDiscountType.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If

            ElseIf Me.CboDiscountType.ListIndex = 1 Then

                If Me.CboTrans.ListIndex <> 1 And Me.CboTrans.ListIndex <> 2 Then
                    Msg = "«·Œ’„ «·„ﬂ ”» ÌﬂÊ‰ „⁄ ( ›« Ê—… «·‘—«¡ «Ê „— Ã⁄ «·„»Ì⁄« )"
                    Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ ‰Ê⁄ «·Œ’„ «Ê ‰Ê⁄ «·›« Ê—… «·„Õœœ…..!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    CboDiscountType.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If
            End If

            If Trim(Me.TxtTransSerial.text) = "" Then
                Msg = "»—Ã«¡ ≈œŒ«· —ﬁ„ «·›« Ê—…..!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Sub
            Else

                If Me.CboTrans.ListIndex = 0 Then '›« Ê—… „»Ì⁄« 
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 2 Or 21)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 1 Then '›« Ê—… „‘ —Ì« 
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 1 Or 22)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 2 Then '„— Ã⁄ „»Ì⁄« 
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 9)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 3 Then '„— Ã⁄ „‘ —Ì« 
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 5)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 4 Then

                    If CheckDebitMaintaince(val(Me.TxtTransSerial.text)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 5 Then
                    Msg = "⁄›Ê« .. Ã«—Ï  ÿÊÌ— «·»—‰«„Ã .. ·⁄„· «·Œ’Ê„«  ··‹ «·Œœ„« "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
        End If
    
        If TxtNoteSerial.text = "" Then
            If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
                End If
            End If
        End If
    
    
        If TxtNoteSerial1.text = "" Then
        If Voucher_coding(val(my_branch), XPDtbTrans.value, 62, mNoteTypeSer) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ Œ’„  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        Else
                       
            If Voucher_coding(val(my_branch), XPDtbTrans.value, 62, mNoteTypeSer) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ    ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            Else
                TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 62, mNoteTypeSer)
            End If
        End If
    End If
    
    
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then
            rs.AddNew
              XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
            rs("NoteID").value = val(XPTxtID.text)

        
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
        
        rs("branch_no").value = val(Me.Dcbranch.BoundText)
        rs("EmpId").value = IIf(Me.DcEmp.BoundText = "", Null, val(Me.DcEmp.BoundText))
        rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '„”·”· «–‰ «·’—›
        rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, val(XPTxtVal.text))
         Me.Lb_note_value_by_characters.Caption = WriteNo(Format(val(Me.XPTxtVal.text) + val(TxtVAT.text), "0.00"), 0, True, ".")
        rs("note_value_by_characters").value = Trim$(Me.Lb_note_value_by_characters.Caption)
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        rs("BankID").value = Null
        rs("VAT").value = val(TxtVAT.text)
        rs("VATYou").value = val(TxtVATYou.text)
        rs("TotalValue").value = val(TxtTotal.text)
        rs("Account_DebitSide").value = IIf(Me.DcboDebitSide.BoundText = "", Null, (Me.DcboDebitSide.BoundText))
        rs("Account_CreditSide").value = IIf(Me.DcboCreditSide.BoundText = "", Null, (Me.DcboCreditSide.BoundText))
        rs("ORDER_NO").value = Trim(txtORDER_NO.text)
        rs("FiterWaiver").value = val(txtFiterWaiver.text)
        rs("FiterWaiverNoteSerial").value = Trim(txtFiterWaiverNoteSerial.text)
        
        
        If Me.CboDiscountType.ListIndex = 0 Then
            'Œ’„ „”„ÊÕ »Â
            rs("NoteType").value = 9
        ElseIf Me.CboDiscountType.ListIndex = 1 Then
            'Œ’„ „ﬂ ”»
            rs("NoteType").value = 10
            
         ElseIf Me.CboDiscountType.ListIndex = 2 Then
            'œÌÊ‰ „⁄œÊ„…
           rs("NoteType").value = 8034
        ElseIf Me.CboDiscountType.ListIndex = 3 Then
            '„œÌ‰
            rs("NoteType").value = 9082
        ElseIf Me.CboDiscountType.ListIndex = 4 Then
          'œ«∆‰
            rs("NoteType").value = 9083
       ElseIf Me.CboDiscountType.ListIndex = 5 Then
          '
            rs("NoteType").value = 9089
       ElseIf Me.CboDiscountType.ListIndex = 6 Then
          '
            rs("NoteType").value = 9090
 ElseIf Me.CboDiscountType.ListIndex = 7 Then
          '
            rs("NoteType").value = 9099
            
        End If


    rs("Currency_id").value = IIf(DcCurrency.BoundText = "", Null, val(DcCurrency.BoundText))
    rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.text), 1, txt_Currency_rate.text)
    rs("DateRec").value = txtDateRec.value
    rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
    rs("CIBAN").value = TXTIban.text
    rs("Invoicetype").value = Me.DefaultInvoicetype.ListIndex


        rs("NoteDate").value = XPDtbTrans.value

        If Me.ChkTrans.value = vbChecked Then
            If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                rs("Transaction_ID").value = val(Me.TxtTransID.text)
                rs("MaintananceID").value = Null
            ElseIf Me.CboTrans.ListIndex = 2 Then
                rs("Transaction_ID").value = Null
                rs("MaintananceID").value = val(Me.TxtTransID.text)
            End If

        Else
            rs("Transaction_ID").value = Null
            rs("MaintananceID").value = Null
        End If

        rs("CashingType").value = IIf(DCboCashType.ListIndex = -1, Null, DCboCashType.ListIndex)
        rs("CusID").value = IIf(DBCboClientName.text = "", Null, val(DBCboClientName.BoundText))
        rs("BoxID").value = Null
        rs("UserID").value = user_id
        rs("numbering_type").value = sand_numbering_type(0) '„”·”· «·ﬁÌœ
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)
    
        rs.update

        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            Set RsDev = New ADODB.Recordset
         '   RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                     StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
            '«·ÿ—› «·„œÌ‰
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 1
            RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
            If val(CboDiscountType.ListIndex) = 3 Then
            RsDev("Value").value = val(Me.XPTxtVal.text) + val(TxtVAT.text)
            Else
            RsDev("Value").value = val(Me.XPTxtVal.text)
            End If
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
           If (val(CboDiscountType.ListIndex) = 4 Or val(CboDiscountType.ListIndex) = 0) And val(Me.TxtVAT.text) > 0 Then
            GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, AccountVATCreit, 1, 13
            If AccountVATCreit = "" Then
                GetValueAddedAccount XPDtbTrans.value, , AccountVATCreit, 1, 23
            End If
            
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 2
            RsDev("Account_Code").value = AccountVATCreit
            RsDev("Value").value = val(Me.TxtVAT.text)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & " «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì… -«‘⁄«— œ«∆‰"
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
            End If
            '«·ÿ—› «·œ«∆‰
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 3
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
           If val(CboDiscountType.ListIndex) = 4 Or val(CboDiscountType.ListIndex) = 5 Or val(CboDiscountType.ListIndex) = 0 Or val(CboDiscountType.ListIndex) = 1 Then
            RsDev("Value").value = val(Me.XPTxtVal.text) + val(TxtVAT)
           Else
            RsDev("Value").value = val(Me.XPTxtVal.text)
          End If
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
            LblDevID.Caption = LngDevID
            lbl(11).Caption = SystemOptions.SysCurrentAccountIntervalID
        End If
           If (val(CboDiscountType.ListIndex) = 5 Or val(CboDiscountType.ListIndex) = 1 Or val(CboDiscountType.ListIndex) = 3) And val(Me.TxtVAT.text) > 0 Then
            GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, mIndexVat
'            GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 18
                 RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 5
            RsDev("Account_Code").value = AccountVATCreit
            RsDev("Value").value = val(Me.TxtVAT.text)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & "«·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì… «‘⁄«— „œÌ‰"
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
            End If
            
            saveBillBuy
            
            If val(txtFiterWaiver) <> 0 Then
              Dim s As String
            
            s = "Update TblOtheExpensAqar  set Discount2 = " & val(TxtTotal) & " where id = " & val(txtFiterWaiver)
            Cn.Execute s
            
            End If
 
  
    
            
            
        Cn.CommitTrans
        BeginTrans = False
        SaveQRCode "Notes", "NoteID", val(XPTxtID), TxtNoteSerial1.text, (XPDtbTrans.value), _
        (TxtTotal.text), Picture1, 0, (TxtVAT.text), (TxtTotal.text)
        
        
        If Me.CboDiscountType.ListIndex = 3 Or Me.CboDiscountType.ListIndex = 4 Then
            If Not chkTaxExempt.value = vbChecked And SystemOptions.ApplyEinvoice Then savenewelectroncic
            End If
        
        If SystemOptions.IsBluee = True And (Me.CboDiscountType.ListIndex = 3 Or Me.CboDiscountType.ListIndex = 4) Then
 
   
                MsgBox SENDEINVOICE(Me.XPTxtID, True, val(Me.DBCboClientName.BoundText), 3, "Notes", "NoteID"), vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  
        End If

        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

        If Me.ChkTrans.value = vbUnchecked Then
            Me.CboTrans.ListIndex = -1
            Me.TxtTransSerial.text = ""
            Me.TxtTransID.text = ""
        End If

        CuurentLogdata
        If mIsNoMsg Then Exit Sub
        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
        
            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End Select

        TxtModFlg.text = "R"
        Retrive
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Function CheckDebitMaintaince(LngTransID As Long) As Boolean
    Dim Msg As String
    Dim RsTemp As ADODB.Recordset
    Dim DblCreditNoteValue As Double
    Dim LngDebitNoteID As Long
    Dim StrSQL As String

    CheckDebitMaintaince = False

    If LngTransID = 0 Then
        Msg = "⁄›Ê« .. ·« ÊÃœ ›« Ê—… »Â–« «·„”·”· „”Ã·… ›Ï «·»—‰«„Ã..!!!"
        Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…..!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtTransSerial.SetFocus
        Exit Function
    ElseIf LngTransID <> 0 Then
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select CusID,PaymentType From TblMaintenece where MaintananceID=" & LngTransID & ""
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If RsTemp("PaymentType").value = 0 Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & CHR(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  Õ’Ì· ·Â« „ﬁ»Ê÷« "
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If LngTransID <> val(Me.TxtTransID.text) Then
                Me.TxtTransID.text = LngTransID
            End If
        
            DblCreditNoteValue = 0
            StrSQL = "SELECT Notes.Note_Value, Notes.NoteID, TblMaintenece.MaintananceID," & "TblMaintenece.PaymentType, TblMaintenece.MType "
            StrSQL = StrSQL + " FROM TblMaintenece INNER JOIN Notes ON " & "TblMaintenece.MaintananceID = Notes.MaintananceID " & " WHERE (((Notes.NoteType)=1)) AND TblMaintenece.MaintananceID=" & LngTransID & ""
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                LngDebitNoteID = RsTemp("NoteID").value
                DblCreditNoteValue = IIf(IsNull(RsTemp("Note_Value").value), 0, RsTemp("Note_Value").value)
                '«· «ﬂœ „‰ «‰ Â–Â «·›« Ê—… ·Ì”  ·Â« √ﬁ”«ÿ
                'ÕÌÀ «‰ «·√ﬁ”«ÿ ·«Ì„ﬂ‰  Õ’Ì·Â« „‰ Â‰«
                StrSQL = "Select * From InstallMent Where NoteID=" & LngDebitNoteID & ""
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly

                If Not (RsTemp.BOF Or RsTemp.EOF) Then
                    If RsTemp.RecordCount > 0 Then
                        Msg = "⁄›Ê« .. «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ﬁœ  „  ﬁ”ÌÿÂ«..!!"
                        Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘… «·„ﬁ»Ê÷« "
                        Msg = Msg & CHR(13) & "≈” Œœ„ ‘«‘…  Õ’Ì· «·√ﬁ”«ÿ »œ·« „‰Â«"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Exit Function
                    End If
                End If

            Else
                'LngDebitNoteID
                Msg = "·«ÌÊÃœ «Ê—«ﬁ „«·Ì… √Ã·… ⁄·Ï Â–Â «·›« Ê—…..!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Function
            End If

            If DblCreditNoteValue < val(Me.XPTxtVal.text) Then
                Msg = "⁄›Ê« ..."
                Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… .. «’€— „‰ «·ﬁÌ„…"
                Msg = Msg & CHR(13) & "«·„—«œ  ”ÃÌ·Â« «·√‰..»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·….!"
                Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.XPTxtVal.SetFocus
                Exit Function
            End If

            Set RsTemp = New ADODB.Recordset
        
            StrSQL = "SELECT  TblMaintenece.MaintananceID," & "TblMaintenece.MType, TblMaintenece.PaymentType," & "Sum(Notes.Note_Value) AS SumNote_Value "
            StrSQL = StrSQL + " FROM TblMaintenece INNER JOIN Notes ON TblMaintenece.MaintananceID =" & "Notes.MaintananceID " & " Where ((Notes.NoteType = 4) And TblMaintenece.MaintananceID = " & LngTransID & ")"

            If Me.TxtModFlg.text = "E" Then
                StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.text & ""
            End If

            StrSQL = StrSQL + " GROUP BY TblMaintenece.MaintananceID," & "TblMaintenece.MType, TblMaintenece.PaymentType"
        
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If DblCreditNoteValue = RsTemp("SumNote_Value").value Then
                    Msg = "⁄›Ê« ...!!!!!"
                    Msg = Msg & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
                    Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «Ì… „ﬁ»Ê÷«  ≈÷«›Ì… ⁄·ÌÂ«."
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Function
                ElseIf RsTemp("SumNote_Value").value + val(Me.XPTxtVal.text) > DblCreditNoteValue Then
                    Msg = "⁄›Ê« ..."
                    Msg = Msg & CHR(13) & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  „”»ﬁ« ·Â–Â «·›« Ê—…"
                    Msg = Msg & CHR(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
                    Msg = Msg & CHR(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
                    Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                    Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                    Msg = Msg & CHR(13) & "ﬁÌ„… «·„ﬁ»Ê÷«  «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").value
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Function
                End If
            End If

        Else
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
            Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtTransSerial.SetFocus
            Exit Function
        End If
    End If

    CheckDebitMaintaince = True
    Exit Function
ErrTrap:
End Function

Private Function CheckDebitTrans(LngTransID As Long) As Boolean
    Dim Msg As String
    Dim RsTemp As ADODB.Recordset
    Dim DblCreditNoteValue As Double
    Dim LngDebitNoteID As Long
    Dim StrSQL As String

    CheckDebitTrans = False

    If LngTransID = 0 Then
        Msg = "⁄›Ê« .. ·« ÊÃœ ›« Ê—… »Â–« «·„”·”· „”Ã·… ›Ï «·»—‰«„Ã..!!!"
        Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…..!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtTransSerial.SetFocus
        Exit Function
    ElseIf LngTransID <> 0 Then
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select CusID,PaymentType From Transactions where Transaction_ID=" & LngTransID & ""
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If RsTemp("PaymentType").value = 0 Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & CHR(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰ ⁄„· Œ’Ê„«  „”„ÊÕ… ⁄·ÌÂ«"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If LngTransID <> val(Me.TxtTransID.text) Then
                Me.TxtTransID.text = LngTransID
            End If
        
            DblCreditNoteValue = 0
            StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType, " & "Notes.Note_Value, Notes.NoteID "
            StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & "Notes.Transaction_ID WHERE (Notes.NoteType=1) AND Transactions.Transaction_ID= " & LngTransID & ""
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                LngDebitNoteID = RsTemp("NoteID").value
                DblCreditNoteValue = IIf(IsNull(RsTemp("Note_Value").value), 0, RsTemp("Note_Value").value)
                '«· «ﬂœ „‰ «‰ Â–Â «·›« Ê—… ·Ì”  ·Â« √ﬁ”«ÿ
                'ÕÌÀ «‰ «·√ﬁ”«ÿ ·«Ì„ﬂ‰  Õ’Ì·Â« „‰ Â‰«
                StrSQL = "Select * From InstallMent Where NoteID=" & LngDebitNoteID & ""
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly

                If Not (RsTemp.BOF Or RsTemp.EOF) Then
                    If RsTemp.RecordCount > 0 Then
                        Msg = "⁄›Ê« .. «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ﬁœ  „  ﬁ”ÌÿÂ«..!!"
                        Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘… «·Œ’Ê„« "
                        Msg = Msg & CHR(13) & "≈” Œœ„ ‘«‘…  Õ’Ì· «·√ﬁ”«ÿ »œ·« „‰Â«"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Exit Function
                    End If
                End If

            Else
                'LngDebitNoteID
                Msg = "·«ÌÊÃœ «Ê—«ﬁ „«·Ì… √Ã·… ⁄·Ï Â–Â «·›« Ê—…..!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Function
            End If

            If DblCreditNoteValue < val(Me.XPTxtVal.text) Then
                Msg = "⁄›Ê« ..."
                Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… .. «’€— „‰ «·ﬁÌ„…"
                Msg = Msg & CHR(13) & "«·„—«œ  ”ÃÌ·Â« «·√‰..»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·….!"
                Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.XPTxtVal.SetFocus
                Exit Function
            End If

            Set RsTemp = New ADODB.Recordset
            StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType," & "Sum(Notes.Note_Value) AS SumNote_Value "
            StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & "Notes.Transaction_ID "

            If Me.CboDiscountType.ListIndex = 0 Then
                '«·„ﬁ»Ê÷«  Ê«·Œ’„ «·„”„ÊÕ »Â
                StrSQL = StrSQL + " Where ((Notes.NoteType = 4 OR Notes.NoteType = 9) "
            ElseIf Me.CboDiscountType.ListIndex = 1 Then
                '«·„œ›Ê⁄«  Ê«·Œ’„ «·„ﬂ ”»
                StrSQL = StrSQL + " Where ((Notes.NoteType = 5 OR Notes.NoteType = 10) "
            End If

            StrSQL = StrSQL + " And Transactions.Transaction_ID = " & LngTransID & ")"
        
            If Me.TxtModFlg.text = "E" Then
                StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.text & ""
            End If

            StrSQL = StrSQL + " GROUP BY Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType "
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If DblCreditNoteValue = RsTemp("SumNote_Value").value Then
                    If Me.CboDiscountType.ListIndex = 0 Then
                        Msg = "⁄›Ê« ...!!!!!" & CHR(13)
                        Msg = Msg & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  √Ê (⁄„· Œ’Ê„«  „”„ÊÕ…) ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
                        Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  ”ÃÌ· √Ì… Œ’Ê„«  ≈÷«›Ì… ⁄·ÌÂ«."
                    ElseIf Me.CboDiscountType.ListIndex = 1 Then
                        Msg = "⁄›Ê« ...!!!!!" & CHR(13)
                        Msg = Msg & "·ﬁœ  „  ”ÃÌ· „œ›Ê⁄«   √Ê (⁄„· Œ’Ê„«  „ﬂ ”»…) ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
                        Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  ”ÃÌ· «Ì… Œ’Ê„«  ≈÷«›Ì… ⁄·ÌÂ«."
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Function
                ElseIf RsTemp("SumNote_Value").value + val(Me.XPTxtVal.text) > DblCreditNoteValue Then

                    If Me.CboDiscountType.ListIndex = 0 Then
                        Msg = "⁄›Ê« ..."
                        Msg = Msg & CHR(13) & "·ﬁœ  „  Õ’Ì· „ﬁ»Ê÷«  √Ê (⁄„· Œ’Ê„«  „”„ÊÕ…) „”»ﬁ« ·Â–Â «·›« Ê—…"
                        Msg = Msg & CHR(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
                        Msg = Msg & CHR(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
                        Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                        Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                        Msg = Msg & CHR(13) & "ﬁÌ„… «·„œ›Ê⁄«  «Ê «·Œ’Ê„«  «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").value
                    ElseIf Me.CboDiscountType.ListIndex = 1 Then
                        Msg = "⁄›Ê« ..."
                        Msg = Msg & CHR(13) & "·ﬁœ  „  ”ÃÌ· „œ›Ê⁄«  √Ê (⁄„· Œ’Ê„«  „ﬂ ”»…) „”»ﬁ« ·Â–Â «·›« Ê—…"
                        Msg = Msg & CHR(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
                        Msg = Msg & CHR(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
                        Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                        Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                        Msg = Msg & CHR(13) & "ﬁÌ„… «·„œ›Ê⁄«  «Ê «·Œ’Ê„«  «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").value
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Function
                End If
            End If

        Else
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.text)
            Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.text
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtTransSerial.SetFocus
            Exit Function
        End If
    End If

    CheckDebitTrans = True
    Exit Function
ErrTrap:
End Function

Private Sub Label29_Click()
Frame12.Visible = False
End Sub

Private Sub txtFiterWaiver_Change()
' If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
'    Dim s As String
'    Dim rsDummy As New ADODB.Recordset
'    s = "Select RenterID from TblFiterWaiver where id = " & val(txtFiterWaiver)
'    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
'    If Not rsDummy.EOF Then
'        DBCboClientName.BoundText = val(rsDummy!RenterID & "")
'    End If
' End If
End Sub

Private Sub txtFiterWaiverNoteSerial_Change()
 If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    s = "Select CusID,Id from TblOtheExpensAqar where NoteSerial1 = '" & Trim(txtFiterWaiverNoteSerial) & "'"
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        DBCboClientName.BoundText = val(rsDummy!CusID & "")
        txtFiterWaiver = val(rsDummy!ID & "")
    End If
 End If
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
        CboDiscountType.Enabled = False
            '        Me.Caption = "«·Œ’Ê„«  «·„”„ÊÕ… Ê«·„ﬂ ”»…"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            XPTxtVal.locked = True
            XPDtbTrans.Enabled = False
            XPMTxtRemarks.locked = True
            DBCboClientName.locked = True
            DCboCashType.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

            'Fra.Enabled = False
            ChkTrans.Enabled = False

        Case "N"
            CboDiscountType.Enabled = True
            '        Me.Caption = "«·Œ’Ê„«  «·„”„ÊÕ… Ê«·„ﬂ ”»…( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
            XPDtbTrans.Enabled = True
            XPTxtVal.locked = False
            XPMTxtRemarks.locked = False
            DBCboClientName.locked = False
            XPDtbTrans.value = Date
            DCboCashType.locked = False
            DCboCashType.ListIndex = 0
            'Fra.Enabled = True
            ChkTrans.Enabled = True

        Case "E"
            '        Me.Caption = "«·Œ’Ê„«  «·„”„ÊÕ… Ê«·„ﬂ ”»…(  ⁄œÌ· )"
            CboDiscountType.Enabled = True
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
            XPDtbTrans.Enabled = True
            XPMTxtRemarks.locked = False
            DBCboClientName.locked = False
            DCboCashType.locked = False
            'Fra.Enabled = True
            ChkTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

 
 Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
    If DCboCashType.ListIndex = 0 Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
    Else
    GetCustomersDetail CUSTID, , TxtSearchCode.text, 2
    End If
        DBCboClientName.BoundText = CUSTID

    End If

End Sub
Sub Calculte(Optional Ind As Integer = 0)
If Me.TxtModFlg.text <> "R" Then
Dim AccountVATCreit As String
Dim Percetage As Double
If Ind = 0 Or val(TxtVATYou.text) <> 0 Then
If val(CboDiscountType.ListIndex) = 0 Then 'Œ’„ „”„ÊÕ

GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 18
'     PercentgValueAddedAccount_Transec XPDtbTrans.value, 18, 1, AccountVATCreit, Percetage
  
ElseIf val(CboDiscountType.ListIndex) = 1 Then ' „ﬂ ”»
'GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 0, 17
'      PercentgValueAddedAccount_Transec XPDtbTrans.value, 17, 0, AccountVATCreit, Percetage
      
ElseIf val(CboDiscountType.ListIndex) = 3 Then
      GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 0, 17
      PercentgValueAddedAccount_Transec XPDtbTrans.value, 17, 0, AccountVATCreit, Percetage
      
ElseIf val(CboDiscountType.ListIndex) = 4 Then GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 18
     PercentgValueAddedAccount_Transec XPDtbTrans.value, 18, 1, AccountVATCreit, Percetage
 ElseIf val(CboDiscountType.ListIndex) = 5 Then
    GetValueAddedAccount XPDtbTrans.value, , AccountVATCreit, 1, 13
    DcboCreditSide.BoundText = AccountVATCreit
 ElseIf val(CboDiscountType.ListIndex) = 6 Then
    GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 14
    DcboDebitSide.BoundText = AccountVATCreit
End If
TxtVATYou.text = Percetage
               ' If Percetage = 0 Then
               ' Percetage = 1
               ' End If
                Percetage = Percetage / 100 + 1
                XPTxtVal.text = val(TxtTotal.text) / Percetage
                TxtVAT.text = val(XPTxtVal.text) * val(TxtVATYou.text) / 100
 ElseIf val(TxtVATYou.text) = 0 Then
                XPTxtVal.text = val(TxtTotal.text)
                TxtVAT.text = 0
 End If
End If
If val(CboDiscountType.ListIndex) = 5 Or val(CboDiscountType.ListIndex) = 6 Then
TxtVATYou.text = 0
XPTxtVal.text = val(TxtTotal.text)
TxtVAT.text = 0
End If
End Sub

Private Sub TxtTotal_Change()
Calculte
XPTxtVal_Change
End Sub

Private Sub TxtTransID_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If Me.TxtTransID.text <> "" Then
            If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Or Me.CboTrans.ListIndex = 2 Or Me.CboTrans.ListIndex = 3 Then
                Me.TxtTransSerial.text = GetTransIDSerial(1, val(Me.TxtTransID.text))
            Else
                Me.TxtTransSerial.text = Me.TxtTransID.text
            End If
        End If
    End If

End Sub

Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTransSerial.text, 1)
End Sub

Private Sub TxtVATYou_Change()
Calculte 1
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Cancel = True
End Sub
Sub RelineBuy()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid1
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           Sm = Sm + val(.TextMatrix(i, .ColIndex("RemainingValue")))
           End If
           Next i
  
    End With
   Label28.Caption = Sm
End Sub
Sub RetriveBillBuy(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
Set Rs8 = New ADODB.Recordset
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.rows = 1
End With
sql = " SELECT      dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
sql = sql & "                      dbo.Transactions.ManualNO, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.CusID,"
sql = sql & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.TotalPayed, dbo.Transactions.OldContID,"
sql = sql & "                      dbo.transactions.OldValue , dbo.transactions.dueDate, dbo.transactions.Vat, dbo.transactions.Transaction_NetValue"
sql = sql & " FROM         dbo.Transactions LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & "  WHERE     (dbo.Transactions.PaymentType = 1) AND (dbo.Transactions.Transaction_Type = 21 OR"
sql = sql & "                       dbo.Transactions.Transaction_Type = 2 or dbo.Transactions.Transaction_Type = 71) AND (dbo.Transactions.TotalPayed IS NULL OR"
sql = sql & "                       dbo.Transactions.TotalPayed = 0) AND (dbo.Transactions.CusID = " & CuID & ")"
sql = sql & "  ORDER BY dbo.Transactions.DueDate ,dbo.Transactions.NoteSerial1"

Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
VSFlexGrid1.Enabled = True


        VSFlexGrid1.Enabled = True
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.rows = 1
    .rows = .rows + Rs8.RecordCount
.rows = .FixedRows + Rs8.RecordCount
Rs8.MoveFirst
For i = .FixedRows To Rs8.RecordCount
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("BranchId").value), 0, Rs8("BranchId").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
Else
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
End If

.TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(Rs8("DueDate").value), "", Rs8("DueDate").value)
.TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs8("Transaction_ID").value), 0, Rs8("Transaction_ID").value)
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("Transaction_Date").value), "", Rs8("Transaction_Date").value)
.TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
.TextMatrix(i, .ColIndex("too")) = IIf(IsNull(Rs8("ManualNO").value), "", Rs8("ManualNO").value)
.TextMatrix(i, .ColIndex("Note_Value")) = val(IIf(IsNull(Rs8("Transaction_NetValue").value), IIf(IsNull(Rs8("OldValue").value), 0, Rs8("OldValue").value), Rs8("Transaction_NetValue").value))
If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteBillBuy(val(.TextMatrix(i, .ColIndex("NoteID"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("Note_Value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
Rs8.MoveNext
Next i
End With
End If
End Sub
Function GeteBillBuy(Optional Transaction_ID As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS Smatiobn"
sql = sql & " From dbo.TblBillBuyPayment2"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
sql = sql & " GROUP BY Transaction_ID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteBillBuy = IIf(IsNull(Rs8("Smatiobn").value), 0, Rs8("Smatiobn").value)
Else
GeteBillBuy = 0
End If
End Function
Public Sub RetriveBillBuyData(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String


   ' On Error GoTo ErrTrap
    Set RsDetails = New ADODB.Recordset
  StrSQL = "   SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblNotesBillBuyPayment2.*"
  StrSQL = StrSQL & "  FROM         dbo.TblNotesBillBuyPayment2 LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblNotesBillBuyPayment2.branch_no = dbo.TblBranchesData.branch_id"
  StrSQL = StrSQL & "  Where (dbo.TblNotesBillBuyPayment2.NoteID1 = " & val(XPTxtID.text) & " and dbo.TblNotesBillBuyPayment2.TransType=2)"
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid1
    .Clear flexClearScrollable, flexClearEverything
    .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To RsDetails.RecordCount
        .TextMatrix(i, .ColIndex("Ser")) = i

            .TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(RsDetails("branch_no").value), 0, RsDetails("branch_no").value)
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
            Else
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_namee").value), 0, RsDetails("branch_namee").value)
            End If
            .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsDetails("NoteID").value), 0, RsDetails("NoteID").value)
            .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDetails("NoteSerial1").value), 0, RsDetails("NoteSerial1").value)
            .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsDetails("Note_Value").value), 0, RsDetails("Note_Value").value)
            .TextMatrix(i, .ColIndex("PayedValue")) = IIf(IsNull(RsDetails("PayedValue").value), 0, RsDetails("PayedValue").value)
            .TextMatrix(i, .ColIndex("TransPayedValue")) = IIf(IsNull(RsDetails("TransPayedValue").value), 0, RsDetails("TransPayedValue").value)
            .TextMatrix(i, .ColIndex("too")) = IIf(IsNull(RsDetails("too").value), "", RsDetails("too").value)
            .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(RsDetails("NetValue").value), 0, RsDetails("NetValue").value)
            .TextMatrix(i, .ColIndex("RemainingValue")) = IIf(IsNull(RsDetails("RemainingValue").value), 0, RsDetails("RemainingValue").value)
            .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(((RsDetails("DueDate").value))), " ", ((RsDetails("DueDate").value)))
            .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(((RsDetails("NoteDate").value))), "", ((RsDetails("NoteDate").value)))
            .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
            RsDetails.MoveNext
        Next i
        

    End If
End With
RelineBuy
    RsDetails.Close
    Set RsDetails = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
End Sub
Sub BillCustomer()
Dim Msg As String
If Me.TxtModFlg.text <> "R" Then
If Me.TxtModFlg.text = "N" Then
RetriveBillBuy val(DBCboClientName.BoundText)
End If
If Me.TxtModFlg.text = "E" And (FlgBillBuy = True Or VSFlexGrid1.rows = 1) Then
RetriveBillBuy val(DBCboClientName.BoundText)
End If
End If
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

Private Sub XPDtbTrans_Change()
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    Calculte
End Sub

Private Sub XPTxtVal_Change()
    Me.Lb_note_value_by_characters.Caption = WriteNo(Format(val(XPTxtVal.text) + val(TxtVAT.text), "0.00"), 0, True, ".")
End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim RsDev As ADODB.Recordset
    Dim i As Integer
    On Error GoTo ErrTrap


    If mIsNoMsg Then
    Set rs = New ADODB.Recordset
    StrSQL = "Select * From Notes Where (NoteType=9 OR NoteType=10 or NoteType=8034 or NoteType=9082 or NoteType=9083 or NoteType=9089 or NoteType=9090 or NoteType=9099)"
StrSQL = StrSQL & "   AND   branch_no in(" & Current_branchSql & ")"

        If SystemOptions.usertype <> UserAdmin Then
'        StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
    
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    End If
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
Dim Dcombos As New ClsDataCombos
    Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.DcboDebitSide.BoundText = IIf(IsNull(rs("Account_DebitSide").value), "", rs("Account_DebitSide").value)
    Me.DcboCreditSide.BoundText = IIf(IsNull(rs("Account_CreditSide").value), "", rs("Account_CreditSide").value)
    
    Me.DcEmp.BoundText = IIf(IsNull(rs("EmpId")), "", rs("EmpId"))
    XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
        Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
        txtORDER_NO.text = IIf(IsNull(rs("ORDER_NO").value), "", rs("ORDER_NO").value)
    XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", Trim(rs("Note_Value").value))
    XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    DCboCashType.ListIndex = IIf(IsNull(rs("CashingType").value), -1, rs("CashingType").value)
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, False
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
    TxtVAT.text = IIf(IsNull(rs("VAT").value), 0, rs("VAT").value)
    TxtVATYou.text = IIf(IsNull(rs("VATYou").value), 0, rs("VATYou").value)
    TxtTotal.text = IIf(IsNull(rs("TotalValue").value), 0, rs("TotalValue").value)
    txtFiterWaiver.text = IIf(IsNull(rs("FiterWaiver").value), "", rs("FiterWaiver").value)
    txtFiterWaiverNoteSerial.text = IIf(IsNull(rs("FiterWaiverNoteSerial").value), "", rs("FiterWaiverNoteSerial").value)
    
    
    If rs("NoteType").value = 9 Then
        Me.CboDiscountType.ListIndex = 0
    ElseIf rs("NoteType").value = 10 Then
        Me.CboDiscountType.ListIndex = 1
    ElseIf rs("NoteType").value = 8034 Then
        Me.CboDiscountType.ListIndex = 2
    ElseIf rs("NoteType").value = 9082 Then
        Me.CboDiscountType.ListIndex = 3
    ElseIf rs("NoteType").value = 9083 Then
        Me.CboDiscountType.ListIndex = 4
    ElseIf rs("NoteType").value = 9089 Then
        Me.CboDiscountType.ListIndex = 5
    ElseIf rs("NoteType").value = 9090 Then
        Me.CboDiscountType.ListIndex = 6
    ElseIf rs("NoteType").value = 9099 Then
        Me.CboDiscountType.ListIndex = 7
                
    Else
        Me.CboDiscountType.ListIndex = -1
    End If




    DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)
    Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
    txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    txtDateRec.value = IIf(IsNull(rs("DateRec").value), Date, (rs("DateRec").value))
    zatcaStatus = IIf(IsNull(rs("zatcaStatus").value), 0, rs("zatcaStatus").value)
    TXTIban.text = IIf(IsNull(rs("CIBAN").value), "", (rs("CIBAN").value))
    
    DefaultInvoicetype.ListIndex = IIf(IsNull(rs("Invoicetype").value), 0, rs("Invoicetype").value)
    'StartDateProje.value = IIf(IsNull(rs("StartDateProje").value), Date, rs("StartDateProje").value)
    Me.Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    'DcDiscountAccount.BoundText = IIf(IsNull(rs("DiscountAccount").value), "", rs("DiscountAccount").value)
'    txtid.text = IIf(IsNull(rs("id").value), 0, (rs("id").value))
'    TxtPreVAT.text = IIf(IsNull(rs("PreVAT").value), 0, rs("PreVAT").value)
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
'    dueDate.value = IIf(IsNull(rs("dueDate").value), Date, rs("dueDate").value)
'    dueDate1.value = IIf(IsNull(rs("dueDate1").value), Date, rs("dueDate1").value)
       txtORDER_NO = IIf(IsNull(rs("OrDer_no").value), "", rs("OrDer_no").value)
'    TXTOrDer_no2 = IIf(IsNull(rs("OrDer_no2").value), "", rs("OrDer_no2").value)
    

    If Not IsNull(rs("Transaction_ID").value) Then
        Me.ChkTrans.value = vbChecked
        Me.ChkTrans.Enabled = True
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select * From Transactions Where Transaction_ID=" & rs("Transaction_ID").value
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            Me.TxtTransID.text = RsTemp("Transaction_ID").value
            Me.TxtTransSerial.text = IIf(IsNull(RsTemp("Transaction_Serial").value), "", RsTemp("Transaction_Serial").value)

            If Not (IsNull(RsTemp("Transaction_Type").value)) Then
                If RsTemp("Transaction_Type").value = 1 Then '›« Ê—… ‘—«¡
                    Me.CboTrans.ListIndex = 1
                ElseIf RsTemp("Transaction_Type").value = 9 Then '„— Ã⁄ „»»Ì⁄« 
                    Me.CboTrans.ListIndex = 2
                ElseIf RsTemp("Transaction_Type").value = 2 Then '›« Ê—… »Ì⁄
                    Me.CboTrans.ListIndex = 0
                ElseIf RsTemp("Transaction_Type").value = 5 Then '„— Ã⁄ ‘—«¡
                    Me.CboTrans.ListIndex = 3
                Else
                    Me.CboTrans.ListIndex = -1
                End If
            End If
        End If

    ElseIf Not IsNull(rs("MaintananceID").value) Then
        Me.ChkTrans.value = vbChecked
        Me.CboTrans.ListIndex = 4
        Me.TxtTransID.text = rs("MaintananceID").value
        Me.TxtTransSerial.text = rs("MaintananceID").value
    Else
        Me.ChkTrans.value = vbUnchecked
        Me.CboTrans.ListIndex = -1
        Me.TxtTransID.text = ""
        Me.TxtTransSerial.text = ""
    End If

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.XPTxtID.text)
        StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lbl(11).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For i = 1 To RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                If Me.DcboDebitSide.text = "" Then
                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                 End If
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                If DcboCreditSide.text = "" Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                 End If
                End If

                RsDev.MoveNext
            Next i

        End If
    End If
RetriveBillBuyData
RelineBuy
    '-----------------------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Function AutoCalculate() As Boolean
Dim i As Integer
Dim NetValu As Double
Dim TempValu As Double
Dim RemainValu As Double
NetValu = val(TxtTotal.text)
With VSFlexGrid1
For i = 1 To .rows - 1
RemainValu = val(.TextMatrix(i, .ColIndex("RemainingValue")))
If NetValu >= RemainValu Then
TempValu = RemainValu
NetValu = NetValu - TempValu
Else
TempValu = NetValu
NetValu = 0
End If
If TempValu > 0 Then
  .TextMatrix(i, .ColIndex("TransPayedValue")) = TempValu
  .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
   End If
Next i
End With
If NetValu <> 0 Then
AutoCalculate = False
Else
AutoCalculate = True
End If
End Function
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
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
Sub DeleteBillBuy()
Dim i As Integer
Dim StrSQL As String
With VSFlexGrid1
 For i = .FixedRows To .rows - 1
 If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
      StrSQL = "Update Transactions Set  TotalPayed=0 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
     End If
     Next i
 End With
End Sub
Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
               
                  DeleteBillBuy
              StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.XPTxtID.text) & " and TransType=2"
              Cn.Execute StrSQL, , adExecuteNoRecords
              StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.text) & " and TransType=2"
              Cn.Execute StrSQL, , adExecuteNoRecords
               rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
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

Private Sub WriteDev()
    Dim Account_Code_dynamic As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

        If Me.CboDiscountType.ListIndex = 0 Or Me.CboDiscountType.ListIndex = 4 Then
             
            Account_Code_dynamic = get_account_code_branch(12, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Exit Sub
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·Œ’„ «·„”„ÊÕ »Â ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Exit Sub
         
                End If
            End If
        
            'Œ’„ „”„ÊÕ »Â
            Me.DcboDebitSide.BoundText = Account_Code_dynamic
            ' Me.DcboDebitSide.BoundText = "a3a5"
            Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
        ElseIf Me.CboDiscountType.ListIndex = 1 Or Me.CboDiscountType.ListIndex = 3 Then
            'Œ’„ „ﬂ ”»
            Account_Code_dynamic = get_account_code_branch(13, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Exit Sub
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·Œ’„ «·„ﬂ ”» »Â ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Exit Sub
         
                End If
            End If
        
            Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
            ' Me.DcboCreditSide.BoundText = "a4a4"
            Me.DcboCreditSide.BoundText = Account_Code_dynamic
      
    
    
    
        ElseIf Me.CboDiscountType.ListIndex = 2 Then
            '   œÌÊ‰ „⁄œÊ„…
            Account_Code_dynamic = get_account_code_branch(97, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Exit Sub
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   œÌÊ‰ „⁄œÊ„…    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Exit Sub
         
                End If
            End If
        
            Me.DcboDebitSide.BoundText = Account_Code_dynamic
            ' Me.DcboCreditSide.BoundText = "a4a4"
            Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
          ElseIf Me.CboDiscountType.ListIndex = 3 Or Me.CboDiscountType.ListIndex = 4 Then
          DcboCreditSide.BoundText = ""
          DcboDebitSide.BoundText = ""
        End If
        
       End If

End Sub




Private Sub DcCurrency_Change()

    If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    If Me.DcCurrency.BoundText <> "" Then
        txt_Currency_rate.text = get_currency_rate(Me.DcCurrency.BoundText)
    Else
        txt_Currency_rate.text = 1
    End If

End Sub

Private Sub DcCurrency_Click(Area As Integer)
    DcCurrency_Change
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
 
    If Me.CboDiscountType.ListIndex = 3 Then
        InvoiceTypeCodeID = 383
    ElseIf Me.CboDiscountType.ListIndex = 4 Then
        InvoiceTypeCodeID = 381
    End If

  End If
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

   rs("DocumentCurrencyCode").value = DcCurrency.text
   rs("TaxCurrencyCode").value = DcCurrency.text
  rs("ActualDeliveryDate").value = txtDateRec.value
 rs("LatestDeliveryDate").value = txtDateRec.value
Dim PaymentMeansCode As String
         
            '10 In cash
            '30 Credit
            '42 Payment to bank account
            '48 Bank card
            '1 Instrument not defined(Free text)
            Dim paymentnote
'        If CboPayMentType.ListIndex = 0 Then ' ‰ﬁœ«
'                  PaymentMeansCode = "10"
'                      paymentnote = "Payment by Cash"
'        ElseIf CboPayMentType.ListIndex = 1 Then ' ¬Ã·
'                 PaymentMeansCode = "30"
'                 paymentnote = "Payment by Credit"
'         ElseIf CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then  '  ÕÊÌ· »‰ﬂÌ
'                    If SystemOptions.AllowSalesMultyPayed = True Then
'                     PaymentMeansCode = "48" 'ﬂ«— 
'                      paymentnote = "Payment by Bank Card"
'                    Else
'                    PaymentMeansCode = "42" '»‰ﬂ Õ”«»
'                    paymentnote = "Payment by Bank Account"
'                    End If
'
'         End If
         PaymentMeansCode = "30"
                 paymentnote = "Payment by Credit"
         rs("PaymentMeansCode").value = PaymentMeansCode
      
rs("paymentnote").value = paymentnote
rs.update
End Function




'====================  VB6: Import Excel ? Notes & Vouchers  ====================
' ÷⁄ Â–« «·ﬂÊœ ›Ì «·›Ê—„ «·–Ì ÌÕ ÊÌ ⁄·Ï «·“— cmdImportExcel


' ===== ÀÊ«»  ﬁ«»·… ·· ⁄œÌ· =====


' ·Ê ⁄‰œﬂ « ’«· ⁄«„ Cn ›Ì „‘—Ê⁄ﬂ° ⁄·¯ﬁ «·”ÿ—Ì‰ œÊ·

Private Function GetCn() As ADODB.Connection
    If Not Cn Is Nothing Then
        If Cn.State = adStateOpen Then
            Set GetCn = Cn: Exit Function
        End If
    End If
    Set Cn = New ADODB.Connection
    ' ===== ⁄œ¯· «·”ÿ— «· «·Ì Õ”» ”Ì—›—ﬂ =====
    ' „À«· ÊÌ‰œÊ“ ·ÊÃÊ‰:
    'Cn.Open "Provider=SQLOLEDB;Data Source=.\SQLEXPRESS;Initial Catalog=Gedah;Integrated Security=SSPI;"
    ' „À«· SQL Login:
    Cn.Open "Provider=SQLOLEDB;Data Source=SERVERNAME;Initial Catalog=Gedah;User ID=sa;Password=yourStrong(!)Password;"
    Set GetCn = Cn
End Function

Private Sub cmdImportExcel_Click()
    On Error GoTo eh

    Dim F As String
    F = GetExcelFilePath()
    If Len(F) = 0 Then Exit Sub

    Dim rs As ADODB.Recordset
    Set rs = LoadExcelToRecordset(F)
    If rs Is Nothing Then
        MsgBox " ⁄–¯— ﬁ—«¡… „·› «·≈ﬂ”Ì·/CSV.  √ﬂœ „‰ «·√⁄„œ… Ê√–Ê‰«  „“Êœ ACE/Jet.", vbExclamation
        Exit Sub
    End If

    Dim BranchID As Long, UserID As Long
    BranchID = branch_id
    UserID = user_id

    Dim okCount As Long, errCount As Long, i As Long
    okCount = 0: errCount = 0: i = 0

    Dim CnX As ADODB.Connection
    Set CnX = GetCn()

    Do Until rs.EOF
        i = i + 1

        Dim InvoiceNo As String, customername As String, NoteTypeText As String
        Dim Mobile As String, VATNO As String
        Dim InvoiceDate As Date, Amount As Currency
        Dim IsDiscount As Boolean
        Dim CusID As Long, NoteID As Long
        
        InvoiceNo = Trim$(NzStr(rs.Fields("InvoiceNo").value))
        If InvoiceNo = "" Then GoTo exitss
        customername = Trim$(NzStr(rs.Fields("CustomerName").value))
        NoteTypeText = Trim$(NzStr(GetFld(rs, "NoteTypeText")))
       ' Mobile = Trim$(NzStr(GetFld(rs, "Mobile")))
       ' VATNO = Trim$(NzStr(GetFld(rs, "VATNO")))
        InvoiceDate = CDate(rs.Fields("InvoiceDate").value)
        Amount = CCur(rs.Fields("Amount").value)
        IsDiscount = (NoteTypeText = "Œ’„")

        If Len(customername) = 0 Or Len(InvoiceNo) = 0 Then
            errCount = errCount + 1
            GoTo NextRow
        End If

        ' 1) Upsert ⁄„Ì· (Ì÷»ÿ parent_account „‰ branches.a8)
        CusID = UpsertCustomer(CnX, customername, Mobile, VATNO, BranchID)
        If CusID = 0 Then
            errCount = errCount + 1
            GoTo NextRow
        End If

        ' 2) ≈‰‘«¡ «·≈‘⁄«— »«·ﬁÌ„ «·«› —«÷Ì… + «·ﬁÌÊœ (a12/a13)
        NoteID = CreateNote_Defaults(CnX, CusID, InvoiceDate, Amount, InvoiceNo, BranchID, UserID, IsDiscount, AR_ACCT_SERIAL)
        If NoteID > 0 Then
            okCount = okCount + 1
        Else
            errCount = errCount + 1
        End If

NextRow:
        rs.MoveNext
        DoEvents
    Loop
exitss:
    MsgBox " „  «·„⁄«·Ã… »‰Ã«Õ." & vbCrLf & _
           "«·„⁄«·Ã… «·‰«ÃÕ…: " & okCount & vbCrLf & _
           "⁄œœ «·’›Ê› «· Ì ÕœÀ  »Â« √Œÿ«¡: " & errCount, vbInformation
    Exit Sub

eh:
    MsgBox "Œÿ√ √À‰«¡ «·«” Ì—«œ: " & Err.Description, vbCritical
End Sub

'========================= «” œ⁄«¡ «·≈Ã—«¡«  «·„Œ“‰… =========================

Private Function UpsertCustomer(ByVal Cn As ADODB.Connection, _
                                ByVal CusName As String, ByVal Mobile As String, ByVal VATNO As String, _
                                ByVal BranchID As Long) As Long
    On Error GoTo eh
    Dim Cmd As ADODB.Command
    Set Cmd = New ADODB.Command
    With Cmd
        .ActiveConnection = Cn
        .CommandText = "dbo.Customer_Upsert_FromExcelRow"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 120

        .Parameters.Append .CreateParameter("@CusName", adVarWChar, adParamInput, 300, CusName)
        .Parameters.Append .CreateParameter("@Cus_mobile", adVarWChar, adParamInput, 100, IIf(Len(Mobile) = 0, Null, Mobile))
        .Parameters.Append .CreateParameter("@VATNO", adVarWChar, adParamInput, 64, IIf(Len(VATNO) = 0, Null, VATNO))
        .Parameters.Append .CreateParameter("@BranchId", adInteger, adParamInput, , BranchID)
        .Parameters.Append .CreateParameter("@CusID", adInteger, adParamOutput)

        .Execute , , adExecuteNoRecords
        UpsertCustomer = NzLng(.Parameters("@CusID").value, 0)
    End With
    Exit Function
eh:
    Debug.Print "UpsertCustomer error: "; Err.Description
    UpsertCustomer = 0
End Function

Private Function CreateNote_Defaults(ByVal Cn As ADODB.Connection, _
                                     ByVal CusID As Long, ByVal NoteDate As Date, ByVal Amount As Currency, _
                                     ByVal InvoiceNo As String, ByVal BranchID As Long, ByVal UserID As Long, _
                                     ByVal IsDiscount As Boolean, ByVal ARSerial As String) As Long
    On Error GoTo eh
    Dim Cmd As ADODB.Command
    Set Cmd = New ADODB.Command
    With Cmd
       With Cmd
    .ActiveConnection = Cn
    .CommandText = "dbo.Note_InsertWithVouchers"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 120

    .Parameters.Append .CreateParameter("@CusID", adInteger, adParamInput, , CusID)
    .Parameters.Append .CreateParameter("@NoteDate", adDBDate, adParamInput, , NoteDate)
    .Parameters.Append .CreateParameter("@Amount", adCurrency, adParamInput, , Amount)
    .Parameters.Append .CreateParameter("@InvoiceNo", adVarWChar, adParamInput, 50, InvoiceNo)
    .Parameters.Append .CreateParameter("@BranchId", adInteger, adParamInput, , BranchID)
    .Parameters.Append .CreateParameter("@UserId", adInteger, adParamInput, , UserID)
    .Parameters.Append .CreateParameter("@IsDiscountNote", adBoolean, adParamInput, , (IsDiscount <> 0))
    ' «»⁄  Â‰« Account_Code ··–„„ (√Ê «·”Ì—Ì«· ﬂ‰’)
    ARSerial = ""
    .Parameters.Append .CreateParameter("@AR_ACCT_INPUT", adVarWChar, adParamInput, 100, ARSerial)
    .Parameters.Append .CreateParameter("@NoteId", adInteger, adParamOutput)

    .Execute , , adExecuteNoRecords
    CreateNote_Defaults = NzLng(.Parameters("@NoteId").value, 0)
End With

    End With
    Exit Function
eh:
    Debug.Print "CreateNote_Defaults error: "; Err.Description
    CreateNote_Defaults = 0
End Function

'============================ › Õ „·› «·≈ﬂ”Ì· ============================

Private Function GetExcelFilePath() As String
    On Error Resume Next
    If ControlExists("CommonDialog1") Then
        With Me.CommonDialog1
            .CancelError = False
            .filter = "Excel/CSV|*.xlsx;*.xls;*.csv|All Files|*.*"
            .ShowOpen
            GetExcelFilePath = .FileName
        End With
    Else
        ' »œÌ· »”Ìÿ ›Ì Õ«·… ⁄œ„ ÊÃÊœ CommonDialog
        GetExcelFilePath = InputBox("√œŒ· «·„”«— «·ﬂ«„· ·„·› Excel/CSV:", "«Œ Ì«— „·›")
    End If
End Function

'================== ﬁ—«¡… «·≈ﬂ”Ì·/CSV ≈·Ï Recordset ⁄»— ADO ==================

Private Function LoadExcelToRecordset(ByVal F As String) As ADODB.Recordset
    On Error GoTo eh

    Dim ext As String: ext = LCase$(mId$(F, InStrRev(F, ".") + 1))
    Dim CnX As ADODB.Connection, rs As ADODB.Recordset, sql As String

    Set CnX = New ADODB.Connection
    If ext = "xlsx" Or ext = "xls" Then
        ' «” Œœ„ ACE (Excel 2007+) √Ê Jet (xls ﬁœÌ„)
        On Error Resume Next
        CnX.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & F & _
                 ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;"""
        If Err.Number <> 0 Then
            Err.Clear
            CnX.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & F & _
                     ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;"""
        End If
        On Error GoTo eh

        Dim sheetName As String
        sheetName = GetFirstSheetName(CnX)
        If Len(sheetName) = 0 Then sheetName = "Sheet1$"
        sql = "SELECT InvoiceNo, InvoiceDate, CustomerName, Amount, NoteTypeText FROM [" & sheetName & "]"
    Else
        ' CSV
        Dim folder As String, FileName As String
        folder = left$(F, InStrRev(F, "\"))
        FileName = mId$(F, InStrRev(F, "\") + 1)
        CnX.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & folder & _
                 ";Extended Properties=""text;HDR=YES;FMT=Delimited;IMEX=1;"""
        sql = "SELECT InvoiceNo, InvoiceDate, CustomerName, Amount, NoteTypeText, Mobile, VATNO FROM [" & FileName & "]"
    End If

    Set rs = New ADODB.Recordset
    rs.Open sql, CnX, adOpenForwardOnly, adLockReadOnly
    Set LoadExcelToRecordset = rs
    Exit Function

eh:
    Debug.Print "LoadExcelToRecordset error: "; Err.Description
    Set LoadExcelToRecordset = Nothing
End Function

Private Function GetFirstSheetName(ByVal Cn As ADODB.Connection) As String
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = Cn.OpenSchema(adSchemaTables)
    Do Until rs.EOF
        If rs!TABLE_TYPE = "TABLE" Then
            GetFirstSheetName = rs!table_name
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
End Function

'=============================== Utilities ===============================

Private Function NzStr(v As Variant, Optional ByVal def As String = "") As String
    If IsNull(v) Then NzStr = def Else NzStr = CStr(v)
End Function

Private Function NzLng(v As Variant, Optional ByVal def As Long = 0) As Long
    If IsNull(v) Or Len(Trim$(v & "")) = 0 Then
        NzLng = def
    Else
        NzLng = CLng(v)
    End If
End Function

Private Function GetFld(rs As ADODB.Recordset, ByVal Name As String) As Variant
    On Error Resume Next
    GetFld = rs.Fields(Name).value
End Function

Private Function ControlExists(ByVal ctrlName As String) As Boolean
    On Error Resume Next
    Dim c As Control
    Set c = Me.Controls(ctrlName)
    ControlExists = (Err.Number = 0)
    Err.Clear
End Function
'===============================================================================


