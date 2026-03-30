VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBankDepositeSearch 
   Appearance      =   0  'Flat
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰  «Ìœ«⁄«  »‰ﬂÌ…"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13065
   Icon            =   "FrmBankDepositeSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4800
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   2520
      Width           =   2145
   End
   Begin VB.Frame Frame4 
      Caption         =   "«·‘—Õ"
      Height          =   1095
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   4200
      Width           =   3855
      Begin VB.TextBox TxtRemarks 
         Alignment       =   1  'Right Justify
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "›Ì Õ«·Â «·«Ìœ«⁄ «·‰ﬁœÌ"
      Height          =   855
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   3360
      Width           =   3855
      Begin MSDataListLib.DataCombo DcboBox 
         Height          =   315
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·Œ“Ì‰…"
         Height          =   315
         Index           =   9
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "›Ì Õ«·Â «Ìœ«⁄ ‘Ìﬂ« "
      Height          =   1935
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   3360
      Width           =   4575
      Begin VB.TextBox TxtBankName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   720
         Width           =   2865
      End
      Begin VB.TextBox TxtChequeNumber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1080
         Width           =   2835
      End
      Begin MSComCtl2.DTPicker DTPDueDate 
         Height          =   345
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   100073473
         CurrentDate     =   38979.743287037
      End
      Begin MSDataListLib.DataCombo DcboChequeBox 
         Height          =   315
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ«›Ÿ… «·‘Ìﬂ« "
         Height          =   315
         Index           =   0
         Left            =   3210
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
         Height          =   285
         Index           =   3
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1575
         Width           =   1185
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·»‰ﬂ"
         Height          =   315
         Index           =   3
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·‘Ìﬂ"
         Height          =   285
         Index           =   16
         Left            =   3540
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.TextBox TXTTo 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   4320
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   7800
      Width           =   1365
   End
   Begin VB.ComboBox DCboCashType 
      Height          =   315
      Left            =   8520
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   2610
      Width           =   2835
   End
   Begin MSDataListLib.DataCombo DcboRevenuesTypes 
      Height          =   315
      Left            =   7560
      TabIndex        =   31
      Top             =   7680
      Visible         =   0   'False
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄Ê«„· »ÕÀ ≈÷«›Ì…"
      ForeColor       =   &H00000080&
      Height          =   1995
      Index           =   1
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   8880
      Width           =   6705
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  Œ«’… »«·»ÕÀ ⁄‰ «·‘Ìﬂ« "
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
         Height          =   645
         Index           =   2
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   600
         Width           =   6465
      End
      Begin VB.ComboBox CboPaymentType 
         Height          =   315
         Left            =   3450
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   270
         Width           =   2205
      End
      Begin VB.CheckBox Chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»ÕÀ „ÿ«»ﬁ"
         Height          =   285
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1620
         Width           =   1275
      End
      Begin VB.CheckBox ChkTrans 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰ Õ”«» ›« Ê—… „⁄Ì‰…(”œ«œ «Ê  Õ’Ì· „‰ Õ”«» ›« Ê—…)"
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   330
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtTransSerial 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3180
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1590
         Width           =   2115
      End
      Begin VB.ComboBox CboTrans 
         Height          =   315
         Left            =   3180
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1260
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
         Height          =   315
         Index           =   2
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   270
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Œ — ‰Ê⁄ «·›« Ê—…"
         Height          =   315
         Index           =   0
         Left            =   5340
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«œŒ· —ﬁ„ «·›« Ê—…"
         Height          =   315
         Index           =   6
         Left            =   5340
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1620
         Width           =   1245
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   3150
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2940
      Width           =   1545
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   20
         ToolTipText     =   "«’€— „‰"
         Top             =   0
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "Ì”«ÊÏ"
         Top             =   0
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "«ﬂ»— „‰"
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì «·› —…"
      ForeColor       =   &H00FF0000&
      Height          =   915
      Index           =   0
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3690
      Width           =   2415
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   345
         Left            =   60
         TabIndex        =   4
         Top             =   180
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   100073473
         CurrentDate     =   38979.743287037
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   345
         Left            =   60
         TabIndex        =   5
         Top             =   540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   195
         Width           =   585
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   10
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   555
         Width           =   585
      End
   End
   Begin VB.TextBox TxtValue 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   4740
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2940
      Width           =   2205
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   2250
      TabIndex        =   6
      Top             =   4710
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
      Left            =   1155
      TabIndex        =   7
      Top             =   4710
      Width           =   1035
      _ExtentX        =   1826
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
      Left            =   60
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4710
      Width           =   945
      _ExtentX        =   1667
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
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2445
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   12885
      _cx             =   22728
      _cy             =   4313
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
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmBankDepositeSearch.frx":038A
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
   Begin MSDataListLib.DataCombo DcboUsers 
      Height          =   315
      Left            =   6600
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboExpensesType 
      Height          =   315
      Left            =   4680
      TabIndex        =   0
      Top             =   7530
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboCustomers 
      Height          =   315
      Left            =   6600
      TabIndex        =   3
      Top             =   7200
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdShowMoreOptions 
      Height          =   375
      Left            =   6630
      TabIndex        =   22
      Top             =   7260
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ „ ﬁœ„..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmBankDepositeSearch.frx":060D
      ColorHoverText  =   12582912
      ButtonToggles   =   1
      DrawFocusRectangle=   0   'False
      RightToLeft     =   -1  'True
      ButtonImageToggled=   "FrmBankDepositeSearch.frx":09A7
      ColorToggledHoverText=   192
   End
   Begin MSDataListLib.DataCombo DcboBankName 
      Height          =   315
      Left            =   8520
      TabIndex        =   38
      Top             =   3000
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ”‰œ «·«Ìœ«⁄"
      Height          =   315
      Index           =   10
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   2640
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·»‰ﬂ «·„Êœ⁄ »Â"
      Height          =   285
      Index           =   15
      Left            =   11550
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   3030
      Width           =   1395
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "›« Ê—… «·„Ê—œ"
      Height          =   315
      Index           =   8
      Left            =   5730
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   7710
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ì—«œ«  «·√Œ—Ï"
      Height          =   405
      Index           =   1
      Left            =   5730
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   8130
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·«Ìœ«⁄"
      Height          =   315
      Index           =   7
      Left            =   11760
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2610
      Width           =   1095
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì· «Ê «·„Ê—œ"
      Height          =   405
      Index           =   5
      Left            =   5700
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   7770
      Width           =   975
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„’—Ê›« "
      Height          =   315
      Index           =   4
      Left            =   5790
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   7890
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»·€"
      Height          =   315
      Index           =   1
      Left            =   7590
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3090
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   315
      Index           =   2
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FrmBankDepositeSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_SearchType As Integer
Dim cSearchDcbo(4) As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            Set rs = New ADODB.Recordset
            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.Rows = 2
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            Else

                With Me.FG
                    .Clear flexClearScrollable, flexClearEverything
                    .Rows = .FixedRows
                    .Rows = .FixedRows + rs.RecordCount
                    rs.MoveFirst

                    For i = .FixedRows To rs.RecordCount
                        .TextMatrix(i, .ColIndex("Serial")) = i
                     
                        .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(rs("TblBanksDepositeId").value), "", rs("TblBanksDepositeId").value)
                        .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                    
                        If Not IsNull(rs("RecordDate").value) Then
                            .TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(rs("RecordDate").value)
                        End If
                    
                        .TextMatrix(i, .ColIndex("NoteValue")) = IIf(IsNull(rs("value").value), "", rs("value").value)

                        If Not IsNull(rs("box_or_bank").value) Then
                            If rs("box_or_bank").value = 0 Then
                                '.TextMatrix(i, .ColIndex("PaymentType")) = IIf(IsNull(rs("box_or_bank").value), "", rs("box_or_bank").value)
                                .TextMatrix(i, .ColIndex("PaymentType")) = " «Ìœ«⁄ ‰›œÌ "
                            Else
                                .TextMatrix(i, .ColIndex("PaymentType")) = "«Ìœ«⁄ ‘Ìﬂ«  "
                            End If
                        End If
                   
                        .TextMatrix(i, .ColIndex("DepositeBank")) = IIf(IsNull(rs("DepositeBankName").value), "", rs("DepositeBankName").value)
                   
                        '    .
                   
                        If rs("box_or_bank").value = 0 Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                            
                                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                            Else
                                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxNameE").value), "", rs("BoxNameE").value)
                            End If

                        Else
                   
                            If SystemOptions.UserInterface = ArabicInterface Then
                            
                                .TextMatrix(i, .ColIndex("ChequeBoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                            Else
                                .TextMatrix(i, .ColIndex("ChequeBoxName")) = IIf(IsNull(rs("BoxNameE").value), "", rs("BoxNameE").value)
                            End If

                        End If
                     
                        .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
                    
                        .TextMatrix(i, .ColIndex("ChqueNum")) = IIf(IsNull(rs("ChequeNo").value), "", rs("ChequeNo").value)

                        If Not IsNull(rs("DueDate").value) Then
                            .TextMatrix(i, .ColIndex("DueDate")) = DisplayDate(rs("DueDate").value)
                        End If

                        .TextMatrix(i, .ColIndex("Notes")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
                    
                        rs.MoveNext
                    Next i

                    .AutoSize 0, .Cols - 1, False
                End With

            End If
        
        Case 1
            clear_all Me
        
            FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 2
        
        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub CmdShowMoreOptions_Click()

    If CmdShowMoreOptions.value = True Then
        Me.Fra(1).Visible = True
        Me.Height = Me.Fra(1).top + Fra(1).Height + 600
    Else
        Me.Fra(1).Visible = False
        Me.Height = Me.Fra(1).top + 600
    
    End If

End Sub

Private Sub DCboCashType_Change()
    Dim Dcombos As ClsDataCombos

    If DCboCashType.ListIndex = 0 Then
        '⁄„Ì· «Ê „Ê—œ
        Me.DcboCustomers.Visible = True
        XPLbl(5).Visible = True
        XPLbl(5).Caption = "«”„ «·⁄„Ì· «Ê «·„Ê—œ"
    
        Me.DcboRevenuesTypes.BoundText = ""
        Me.DcboRevenuesTypes.Visible = False
        Me.lbl(1).Visible = False
        Set Dcombos = New ClsDataCombos
        Dcombos.GetCustomersSuppliers 0, Me.DcboCustomers, False
        Set cSearchDcbo(2) = New clsDCboSearch
        Set cSearchDcbo(2).Client = Me.DcboCustomers
    
    ElseIf DCboCashType.ListIndex = 1 Then
        '„ ⁄·ﬁ« 
        Me.DcboRevenuesTypes.BoundText = ""
        Me.DcboRevenuesTypes.Visible = False
        Me.lbl(1).Visible = False
    
        XPLbl(5).Visible = True
        Me.DcboCustomers.Visible = True
    
        XPLbl(5).Caption = "«”„ «·‘Œ’"
        Set Dcombos = New ClsDataCombos
        Dcombos.GetPersons Me.DcboCustomers
        Set cSearchDcbo(2) = New clsDCboSearch
        Set cSearchDcbo(2).Client = Me.DcboCustomers

    ElseIf DCboCashType.ListIndex = 2 Then
    
        '«·≈Ì—«œ«  «·√Œ—Ï
        If Me.SearchType = 4 Then
            Me.DcboRevenuesTypes.Visible = True
            Me.lbl(1).Visible = True
            Me.DcboCustomers.Visible = False
            Me.XPLbl(5).Visible = False
        Else
            Me.DcboRevenuesTypes.Visible = False
            Me.lbl(1).Visible = False
        
            Me.DcboCustomers.Visible = False
            Me.XPLbl(5).Visible = False
        End If

    Else
    End If

End Sub

Private Sub DCboCashType_Click()
    DCboCashType_Change
End Sub

Private Sub Fg_Click()

    With Me.FG

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("NoteID"))) = 0 Then
            Exit Sub
        End If
    
        FrmBankDeposite.Retrive val(.TextMatrix(.Row, .ColIndex("NoteID")))
    
    End With

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim GrdBack As New ClsBackGroundPic

    Set Dcombos = New ClsDataCombos
    CenterForm Me

    FormPostion Me, GetPostion
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DcboUsers
    Dcombos.GetCustomersSuppliers 0, Me.DcboCustomers
    Dcombos.GetExpensesType DcboExpensesType
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetRevenuesTypes Me.DcboRevenuesTypes
    Dcombos.GetChequeBox Me.DcboChequeBox

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DcboBox
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboUsers
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DcboCustomers
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DcboExpensesType

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    With Me.DCboCashType
        .Clear
        .AddItem "«Ìœ«⁄ ‰ﬁœÌ"
        .AddItem " «Ìœ«⁄ ‘Ìﬂ« "
        .AddItem "«·ﬂ·"
        .ListIndex = 2
    End With

    With Me.FG
        Set .WallPaper = GrdBack.SearchWallpaper
        .AutoSize 0, .Cols - 1, False
    End With

    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DTPTo
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

End Sub

Private Sub ChangeLang()
 
    With Me.CboPaymentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheqye"
        .AddItem "All"
        .ListIndex = 2
    End With
 
    Me.Caption = "Search "
 
    XPLbl(3).Caption = "VCHR#"
    XPLbl(0).Caption = "Box"
    XPLbl(1).Caption = "Value"
    XPLbl(2).Caption = "User"
    XPLbl(5).Caption = "Customer"
    XPLbl(8).Caption = "Vendor B#"
    Frame2(0).Caption = "Period"

    lbl(11).Caption = "From"
    lbl(10).Caption = "To"

    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Delete"
    Cmd(2).Caption = "Exit"
 
    With FG
        .TextMatrix(0, .ColIndex("NoteSerial")) = "Vchr#"
        .TextMatrix(0, .ColIndex("NoteDate")) = " Date"
        .TextMatrix(0, .ColIndex("NoteValue")) = "Value "
        .TextMatrix(0, .ColIndex("paymenttype")) = "Payment Type"
        .TextMatrix(0, .ColIndex("CustName")) = "Cust. Name"
        .TextMatrix(0, .ColIndex("NoteCashingType")) = "CashingType"
        .TextMatrix(0, .ColIndex("BoxName")) = "Box Name"
        .TextMatrix(0, .ColIndex("BankName")) = "Bank Name"
  
        .TextMatrix(0, .ColIndex("ChqueNum")) = "ChqueNum"
        .TextMatrix(0, .ColIndex("Notes")) = "Remarks"
        .TextMatrix(0, .ColIndex("UserName")) = "UserName"
  
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    FormPostion Me, SavePostion
    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Erase cSearchDcbo
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 0)
End Sub

Public Property Get SearchType() As Integer
    SearchType = m_SearchType
End Property

Public Property Let SearchType(ByVal vNewValue As Integer)
    m_SearchType = vNewValue

    With Me.FG

        If m_SearchType = 3 Then
            '3 «·»ÕÀ ⁄‰ «·„’—Ê›« 
            .ColHidden(.ColIndex("PaymentType")) = False
            .ColHidden(.ColIndex("CustName")) = True
        
            .ColHidden(.ColIndex("NoteCashingType")) = False
            .ColHidden(.ColIndex("BankName")) = True
            .ColHidden(.ColIndex("ChqueNum")) = True
        
            Me.Caption = "«·»ÕÀ ⁄‰ «·„’—Ê›« "
            Me.XPLbl(4).Visible = True
            Me.DcboExpensesType.Visible = True
        
            Me.XPLbl(5).Visible = False
            Me.DcboCustomers.Visible = False
            Me.CmdShowMoreOptions.value = False
            Me.CmdShowMoreOptions.Visible = False
            Me.DCboCashType.Visible = False
            Me.XPLbl(7).Visible = False
        
            Me.DcboRevenuesTypes.Visible = False
            Me.lbl(1).Visible = False
        ElseIf m_SearchType = 4 Then
            '4 «·»ÕÀ ⁄‰ «·„ﬁ»Ê÷« 
            .ColHidden(.ColIndex("PaymentType")) = True
            .ColHidden(.ColIndex("CustName")) = False
            Me.Caption = "«·»ÕÀ ⁄‰ «·„ﬁ»Ê÷« "
            Me.XPLbl(4).Visible = False
            Me.DcboExpensesType.Visible = False
        
            Me.XPLbl(5).Visible = True
            Me.DcboCustomers.Visible = True
            Me.CmdShowMoreOptions.Visible = True
            Me.CmdShowMoreOptions.value = False
        
            Me.DCboCashType.Visible = True
            Me.XPLbl(7).Visible = True
            Me.XPLbl(7).Caption = "‰Ê⁄ «·„ﬁ»Ê÷« "

            With Me.DCboCashType
                .Clear
                .AddItem " ‰ﬁœÌ"
                .AddItem " ‘Ìﬂ«  "
                .AddItem "«·ﬂ·"
            End With

            DCboCashType.ListIndex = 3

            With Me.CboTrans
                .Clear
                .AddItem "›« Ê—… „»Ì⁄« "
                .AddItem "„— Ã⁄ „‘ —Ì« "
                '.AddItem "’Ì«‰…"
                '.AddItem "Œœ„« "
            End With

        ElseIf m_SearchType = 5 Then
            '5 «·»ÕÀ «·„œ›Ê⁄« 
            Me.Caption = "«·»ÕÀ ⁄‰ «·„œ›Ê⁄« "
            .ColHidden(.ColIndex("PaymentType")) = True
            .ColHidden(.ColIndex("CustName")) = False
            Me.XPLbl(4).Visible = False
            Me.DcboExpensesType.Visible = False
        
            Me.XPLbl(5).Visible = True
            Me.DcboCustomers.Visible = True
            Me.CmdShowMoreOptions.Visible = True
            Me.CmdShowMoreOptions.value = False
        
            Me.DCboCashType.Visible = True
            Me.XPLbl(7).Visible = True
            Me.XPLbl(7).Caption = "‰Ê⁄ «·„œ›Ê⁄« "

            With Me.DCboCashType
                .Clear
                .AddItem "„‰ ⁄„Ì· «Ê „Ê—œ"
                .AddItem "«·„ ⁄·ﬁ« "
                .AddItem "«·ﬂ·"
            End With

            DCboCashType.ListIndex = 2

            With Me.CboTrans
                .Clear
                .AddItem "›« Ê—… „‘ —Ì« "
                .AddItem "„— Ã⁄ „»Ì⁄« "
            End With
        
        End If

    End With

End Property

Private Function Build_Sql() As String
    Dim StrSQL As String
    Dim StrWhere As String
    Dim IntNoteType As Integer

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then

        ' StrSQL = " SELECT     dbo.TblBanksDepositeDetails.TblBanksDepositeId, dbo.TblBanksDepositeDetails.box_or_bank, dbo.TblBanksDepositeDetails.[value], "
        'StrSQL = StrSQL & " dbo.TblBanksDepositeDetails.ChequeNo, dbo.TblBanksDepositeDetails.Remarks, dbo.TblBanksDepositeDetails.BoxID, dbo.TblBoxesData.BoxName,"
        'StrSQL = StrSQL & " dbo.TblBoxesData.BoxNameE, dbo.TblBanksDepositeDetails.bankid, dbo.TblBanksDepositeDetails.BankName, dbo.TblBanksDepositeDetails.DueDate,"
        'StrSQL = StrSQL & " dbo.TblBanksDeposite.NoteSerial1, dbo.TblBanksDeposite.NoteSerial, dbo.TblBanksDeposite.RecordDate, dbo.TblBanksDeposite.bankid AS DepositBank"
        'StrSQL = StrSQL & " FROM         dbo.TblBanksDepositeDetails INNER JOIN"
        'StrSQL = StrSQL & " dbo.TblBoxesData ON dbo.TblBanksDepositeDetails.BoxID = dbo.TblBoxesData.BoxID INNER JOIN"
        'StrSQL = StrSQL & "  dbo.TblBanksDeposite ON dbo.TblBanksDepositeDetails.TblBanksDepositeId = dbo.TblBanksDeposite.id"
        'StrSQL = StrSQL & " WHERE     (1 = 1)"

        StrSQL = " SELECT     dbo.TblBanksDepositeDetails.TblBanksDepositeId, dbo.TblBanksDepositeDetails.box_or_bank, dbo.TblBanksDepositeDetails.[value], "
        StrSQL = StrSQL & " dbo.TblBanksDepositeDetails.ChequeNo, dbo.TblBanksDepositeDetails.Remarks, dbo.TblBanksDepositeDetails.BoxID, dbo.TblBoxesData.BoxName,"
        StrSQL = StrSQL & " dbo.TblBoxesData.BoxNameE, dbo.TblBanksDepositeDetails.BankName, dbo.TblBanksDepositeDetails.DueDate, dbo.TblBanksDeposite.NoteSerial1,"
        StrSQL = StrSQL & " dbo.TblBanksDeposite.NoteSerial, dbo.TblBanksDeposite.RecordDate, dbo.TblBanksDeposite.bankid, dbo.BanksData.BankName AS DepositeBankName"
        StrSQL = StrSQL & " FROM         dbo.TblBanksDepositeDetails INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBoxesData ON dbo.TblBanksDepositeDetails.BoxID = dbo.TblBoxesData.BoxID INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBanksDeposite ON dbo.TblBanksDepositeDetails.TblBanksDepositeId = dbo.TblBanksDeposite.id LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.BanksData ON dbo.TblBanksDeposite.bankid = dbo.BanksData.BankID"
        StrSQL = StrSQL & "  WHERE     (1 = 1)"
    
        'Debug.Print Replace(StrSQL, "dbo.", "", , , vbTextCompare)
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
    End If

    If DCboCashType.ListIndex = 0 Then ' ‰Ê⁄ «·«Ìœ«⁄
 
        StrWhere = StrWhere + " AND  (box_or_bank=0)"
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
 
        StrWhere = StrWhere + " AND  (box_or_bank=1)"
    End If

    If val(Me.TxtValue.text) > 0 Then
        If Me.Opt(1).value = True Then
               
            StrWhere = StrWhere + " AND [value] =" & val(Me.TxtValue.text) & ""
              
        ElseIf Me.Opt(0).value = True Then
              
            StrWhere = StrWhere + " AND [value] >" & val(Me.TxtValue.text) & ""
           
        Else
            
            StrWhere = StrWhere + " AND [value] <" & val(Me.TxtValue.text) & ""
           
        End If
    End If

    If Me.DcboBankName.BoundText <> "" Then
       
        StrWhere = StrWhere + " AND TblBanksDeposite.bankid =" & val(Me.DcboBankName.BoundText) & ""
         
    End If

    If Trim(Me.TxtSerial.text) <> "" Then
        
        StrWhere = StrWhere + " AND NoteSerial1     like'%" & (Me.TxtSerial.text) & "%'"
        
    End If

    If Trim(Me.TXTBankName.text) <> "" Then
        
        StrWhere = StrWhere + " AND BankName like'%" & (Me.TXTBankName.text) & "%'"
        
    End If

    If Trim(Me.TxtChequeNumber.text) <> "" Then
        
        StrWhere = StrWhere + " AND ChequeNo like'%" & (Me.TxtChequeNumber.text) & "%'"
        
    End If

    If Not IsNull(Me.DTPDueDate.value) Then '   «—ÌŒ «·«” Õﬁ«ﬁ
        StrWhere = StrWhere + " AND  DueDate =" & SQLDate(Me.DTPDueDate.value, True) & ""
    End If

    If Me.DcboChequeBox.BoundText <> "" Then
       
        StrWhere = StrWhere + " AND   TblBanksDepositeDetails.BoxID =" & val(Me.DcboChequeBox.BoundText) & ""
         
    End If

    If Me.DcboBox.BoundText <> "" Then
       
        StrWhere = StrWhere + " AND   TblBanksDepositeDetails.BoxID =" & Me.DcboBox.BoundText & ""
         
    End If

    If Trim(Me.TxtRemarks.text) <> "" Then
        
        StrWhere = StrWhere + " AND TblBanksDepositeDetails.Remarks like'%" & (Me.TxtRemarks.text) & "%'"
        
    End If

    '**********************************************************************************************************
    If Not IsNull(Me.DTPFrom.value) Then '  «· «—ÌŒ
        StrWhere = StrWhere + " AND  RecordDate>=" & SQLDate(Me.DTPFrom.value, True) & ""
    End If

    If Not IsNull(Me.DTPTo.value) Then
        StrWhere = StrWhere + " AND  RecordDate <=" & SQLDate(Me.DTPTo.value, True) & ""
    End If
   
    'If Me.DcboUsers.BoundText <> "" Then
           
    '               StrWhere = StrWhere + " AND notes_all.UserID=" & Me.DcboUsers.BoundText & ""
    'End If
 
    StrSQL = StrSQL + StrWhere + " Order By NoteSerial1"
    Build_Sql = StrSQL
End Function

