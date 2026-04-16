VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmNotesSearch 
   Appearance      =   0  'Flat
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ «·„⁄«„·«  «·„«·Ì…"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16335
   Icon            =   "FrmNotesSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   16335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   78
      Top             =   5760
      Width           =   1065
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9750
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   5040
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox TxtAccount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   6510
      Width           =   1305
   End
   Begin VB.TextBox TxtManulaNO 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Frame Frame6 
      Caption         =   "—Ê« » ⁄‰"
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   4680
      Visible         =   0   'False
      Width           =   3255
      Begin VB.ComboBox CmbMonth1 
         Height          =   315
         Left            =   1680
         TabIndex        =   61
         Text            =   "CmbMonth1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox CboYear1 
         Height          =   315
         Left            =   120
         TabIndex        =   60
         Text            =   "CboYear1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "‘Â—"
         Height          =   255
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "”‰…"
         Height          =   255
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox TxtOrderSuppler 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   120
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox TxtDue 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   120
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtEndService 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   120
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txtorder 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3330
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   4680
      Width           =   855
   End
   Begin VB.ComboBox DCboCashType2 
      Height          =   315
      Left            =   2580
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   4320
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   495
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   2760
      Width           =   3975
      Begin VB.TextBox txtRemark 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   120
         Width           =   3045
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„” ›Ìœ"
         Height          =   315
         Index           =   3
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   120
         Width           =   585
      End
   End
   Begin VB.TextBox TXTTo 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5400
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   4320
      Width           =   1305
   End
   Begin VB.ComboBox DCboCashType 
      Height          =   315
      Left            =   0
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   2970
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄Ê«„· »ÕÀ ≈÷«›Ì…"
      ForeColor       =   &H00000080&
      Height          =   1995
      Index           =   1
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   7290
      Width           =   7905
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
         Left            =   1140
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   600
         Width           =   6465
         Begin VB.TextBox TxtChequeNumber 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   2115
         End
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   3000
            TabIndex        =   38
            Top             =   240
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   285
            Index           =   16
            Left            =   2220
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·»‰ﬂ"
            Height          =   285
            Index           =   15
            Left            =   5670
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   270
            Width           =   675
         End
      End
      Begin VB.ComboBox CboPaymentType 
         Height          =   315
         Left            =   4410
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   270
         Width           =   2205
      End
      Begin VB.CheckBox Chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»ÕÀ „ÿ«»ﬁ"
         Height          =   285
         Left            =   2370
         RightToLeft     =   -1  'True
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   330
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtTransSerial 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   4260
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1590
         Width           =   2115
      End
      Begin VB.ComboBox CboTrans 
         Height          =   315
         Left            =   4260
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1260
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
         Height          =   315
         Index           =   2
         Left            =   6630
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   270
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Œ — ‰Ê⁄ «·›« Ê—…"
         Height          =   315
         Index           =   0
         Left            =   6420
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«œŒ· —ﬁ„ «·›« Ê—…"
         Height          =   315
         Index           =   6
         Left            =   6420
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1620
         Width           =   1245
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3180
      Width           =   2265
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«ﬁ·"
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
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "«’€— „‰"
         Top             =   0
         Width           =   915
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
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "Ì”«ÊÏ"
         Top             =   0
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«ﬂ»—"
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
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         ToolTipText     =   "«ﬂ»— „‰"
         Top             =   0
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì «·› —…"
      ForeColor       =   &H00FF0000&
      Height          =   915
      Index           =   0
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3330
      Width           =   2415
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   345
         Left            =   60
         TabIndex        =   6
         Top             =   180
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   241041409
         CurrentDate     =   38979.743287037
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   345
         Left            =   60
         TabIndex        =   7
         Top             =   540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   241041409
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   555
         Width           =   585
      End
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4800
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2490
      Width           =   1905
   End
   Begin VB.TextBox TxtValue 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   5820
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3240
      Width           =   885
   End
   Begin ImpulseButton.ISButton Cmd 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2130
      TabIndex        =   8
      Top             =   6900
      Width           =   1035
      _ExtentX        =   1826
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
      Left            =   1035
      TabIndex        =   9
      Top             =   6900
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6900
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
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   4020
      TabIndex        =   2
      Top             =   2850
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2445
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   16245
      _cx             =   28654
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
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmNotesSearch.frx":038A
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
      Left            =   2460
      TabIndex        =   4
      Top             =   3570
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboExpensesType 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2490
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboCustomers 
      Height          =   315
      Left            =   2460
      TabIndex        =   5
      Top             =   3930
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdShowMoreOptions 
      Height          =   615
      Left            =   3990
      TabIndex        =   26
      Top             =   2340
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1085
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
      ButtonImage     =   "FrmNotesSearch.frx":0689
      ColorHoverText  =   12582912
      ButtonToggles   =   1
      DrawFocusRectangle=   0   'False
      RightToLeft     =   -1  'True
      ButtonImageToggled=   "FrmNotesSearch.frx":0A23
      ColorToggledHoverText=   192
   End
   Begin MSDataListLib.DataCombo DcboRevenuesTypes 
      Height          =   315
      Left            =   2580
      TabIndex        =   48
      Top             =   3930
      Visible         =   0   'False
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcbAqarType 
      Height          =   315
      Left            =   3330
      TabIndex        =   65
      Top             =   5040
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbUnitType 
      Height          =   315
      Left            =   3330
      TabIndex        =   67
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   6150
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbUnitNo 
      Height          =   315
      Left            =   120
      TabIndex        =   69
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   6150
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbAccount 
      Height          =   315
      Left            =   120
      TabIndex        =   71
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   6510
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcsupplier 
      Height          =   315
      Left            =   3300
      TabIndex        =   76
      Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
      Top             =   5370
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbAccount2 
      Height          =   315
      Left            =   360
      TabIndex        =   79
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   5760
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·„ÊŸ›"
      Height          =   195
      Index           =   3
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   80
      Top             =   5760
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   " «·„«·ﬂ"
      Height          =   165
      Index           =   2
      Left            =   7230
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   5370
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·Õ”«»"
      Height          =   195
      Index           =   0
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   72
      Top             =   6510
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÊÕœ…"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   6150
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «·ÊÕœ…"
      Height          =   195
      Index           =   15
      Left            =   6825
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   6150
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·⁄ﬁ«—"
      Height          =   195
      Index           =   1
      Left            =   7545
      TabIndex        =   66
      Top             =   5040
      Width           =   390
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
      Height          =   285
      Index           =   48
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   4680
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ’—› „ ⁄ÂœÌ‰"
      Height          =   195
      Index           =   55
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   4320
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·„” Õﬁ« "
      Height          =   285
      Index           =   54
      Left            =   1110
      TabIndex        =   56
      Top             =   4350
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ‰Â«Ì… «·Œœ„…"
      Height          =   285
      Index           =   70
      Left            =   1110
      TabIndex        =   54
      Top             =   4350
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ÿ·» «·’—›"
      Height          =   285
      Index           =   46
      Left            =   4200
      TabIndex        =   52
      Top             =   4680
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„œ›Ê⁄« "
      Height          =   285
      Index           =   4
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   4350
      Width           =   1125
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "›« Ê—… «·„Ê—œ"
      Height          =   315
      Index           =   8
      Left            =   6930
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   4350
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ì—«œ«  «·√Œ—Ï"
      Height          =   285
      Index           =   1
      Left            =   6810
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3930
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„ﬁ»Ê÷«  «Ê «·„œ›Ê⁄« "
      Height          =   315
      Index           =   7
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   2970
      Visible         =   0   'False
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
      TabIndex        =   20
      Top             =   3930
      Width           =   975
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„’—Ê›« "
      Height          =   315
      Index           =   4
      Left            =   3270
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2490
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”‰œ"
      Height          =   315
      Index           =   3
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2490
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   315
      Index           =   0
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2850
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»·€"
      Height          =   315
      Index           =   1
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3210
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   315
      Index           =   2
      Left            =   6900
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3570
      Width           =   975
   End
End
Attribute VB_Name = "FrmNotesSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_SearchType As Integer
Public m_SearchType2 As Integer
Public person As String
Dim cSearchDcbo(4) As clsDCboSearch
Dim Dcombos As ClsDataCombos
Private Sub CboPayMentType_Change()

    If Me.CboPayMentType.ListIndex = 1 Then
        Fra(2).Enabled = True
        lbl(15).Enabled = True
        Me.DcboBankName.Enabled = True
        lbl(16).Enabled = True
        TxtChequeNumber.Enabled = True
    Else
        Fra(2).Enabled = False
        lbl(15).Enabled = False
        Me.DcboBankName.Enabled = False
        lbl(16).Enabled = False
        TxtChequeNumber.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim IntNoteType As Integer
    IntNoteType = Me.SearchType
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            Set rs = New ADODB.Recordset
        
            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.rows = 2
                If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                Else
                Msg = " No Result"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            Else

                With Me.FG
                    .Clear flexClearScrollable, flexClearEverything
                    .rows = .FixedRows
                    .rows = .FixedRows + rs.RecordCount
                    rs.MoveFirst

                    For i = .FixedRows To rs.RecordCount
                        .TextMatrix(i, .ColIndex("Serial")) = i

                        If Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Then
                            .TextMatrix(i, .ColIndex("A_NoteID")) = IIf(IsNull(rs("A_NoteID").value), "", rs("A_NoteID").value)
                            .TextMatrix(i, .ColIndex("Too")) = IIf(IsNull(rs("too").value), "", rs("too").value)
                    
                        End If
                  
                        .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
                        .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)

                        If Not IsNull(rs("NoteDate").value) Then
                            .TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(rs("NoteDate").value)
                        End If
                        If Me.SearchType = 4 Then
                       .TextMatrix(i, .ColIndex("ManulaNO")) = IIf(IsNull(rs("ManulaNO").value), "", rs("ManulaNO").value)
                        End If
                        If Me.SearchType <> 333 And Me.SearchType <> 360 Then
                        .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                        End If
                        .TextMatrix(i, .ColIndex("PaymentType")) = IIf(IsNull(rs("CashingType").value), "", rs("CashingType").value)
                        If Not IsNull(rs("CashingType").value) Then
                        If rs("CashingType").value = 9 Then
                        .TextMatrix(i, .ColIndex("CustName")) = IIf(IsNull(rs("renterName").value), "", rs("renterName").value)
                        Else
                        .TextMatrix(i, .ColIndex("CustName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                        End If
                        Else
                        If Me.SearchType = 333 Or Me.SearchType = 360 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("Account")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                        Else
                        .TextMatrix(i, .ColIndex("Account")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                        End If
                        .TextMatrix(i, .ColIndex("CustName")) = IIf(IsNull(rs("aqarname").value), "", rs("aqarname").value)
                        Else
                        .TextMatrix(i, .ColIndex("CustName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                        End If
                        End If
                        
                        If Me.SearchType = 333 Or Me.SearchType = 360 Then
                        .TextMatrix(i, .ColIndex("NoteValue")) = IIf(IsNull(rs("value").value), "", rs("value").value)
                          ElseIf Me.SearchType = 7 Or Me.SearchType = 6 Then
                        .TextMatrix(i, .ColIndex("NoteValue")) = IIf(IsNull(rs("Note_Value2").value), "", rs("Note_Value2").value)
                        Else
                        .TextMatrix(i, .ColIndex("NoteValue")) = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
                        End If
                        .TextMatrix(i, .ColIndex("Notes")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
                        .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                        .TextMatrix(i, .ColIndex("PaymentType")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                        If IntNoteType <> 360 And IntNoteType <> 3 And IntNoteType <> 333 And IntNoteType <> 8008 And IntNoteType <> 80 And IntNoteType <> 350 And IntNoteType <> 8063 And IntNoteType <> 300 And IntNoteType <> 3003 And IntNoteType <> 30033 Then
                         .TextMatrix(i, .ColIndex("unittype")) = IIf(IsNull(rs("unittype").value), "", rs("unittype").value)
                          .TextMatrix(i, .ColIndex("UnitNo")) = IIf(IsNull(rs("UnitNo").value), "", rs("UnitNo").value)
                           .TextMatrix(i, .ColIndex("akarid")) = IIf(IsNull(rs("akarid").value), "", rs("akarid").value)
                           End If
                         
                          '  .TextMatrix(i, .ColIndex("PaymentType")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                    
                        If Not IsNull(rs("NoteCashingType").value) Then
                            If rs("NoteCashingType").value = 0 Then
                                .TextMatrix(i, .ColIndex("NoteCashingType")) = "‰ﬁœÌ"
                            ElseIf rs("NoteCashingType").value = 1 Then
                                .TextMatrix(i, .ColIndex("NoteCashingType")) = "‘Ìﬂ"
                            End If

                        Else
                            .TextMatrix(i, .ColIndex("NoteCashingType")) = ""
                        End If

                        .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
                        .TextMatrix(i, .ColIndex("ChqueNum")) = IIf(IsNull(rs("ChqueNum").value), "", rs("ChqueNum").value)
                    
                        rs.MoveNext
                    Next i

                    .AutoSize 0, .Cols - 1, False
                End With

            End If
        
        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 2
opt(1).value = True
        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
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

Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub

Private Sub DCboCashType_Change()
    

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

Private Sub DCboCashType2_Change()
Frame6.Visible = False
TxtOrderSuppler.Visible = False
lbl(55).Visible = False
TxtDue.Visible = False
lbl(54).Visible = False
TxtEndService.Visible = False
lbl(70).Visible = False
Select Case DCboCashType2.ListIndex
Case 9
TxtOrderSuppler.Visible = True
lbl(55).Visible = True
Case 8
TxtDue.Visible = True
lbl(54).Visible = True
Case 10
TxtEndService.Visible = True
lbl(70).Visible = True
Case 6
Frame6.Visible = True
End Select
End Sub
Private Sub YearMonth1()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth1.Clear

    For i = 1 To 12
        CmbMonth1.AddItem MonthName(i)
    Next

    CmbMonth1.ListIndex = Month(Date) - 1
    CboYear1.Clear

    For i = 2010 To 2050
        CboYear1.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear1.NewIndex
        End If

    Next

    CboYear1.ListIndex = IntDefIndex
End Sub
Private Sub DCboCashType2_Click()
DCboCashType2_Change
End Sub

Private Sub fg_Click()

    With Me.FG

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("NoteID"))) = 0 Then
            Exit Sub
        End If

        If Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 6 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 8063 Or Me.SearchType = 360 Then
    
            mdifrmmain.ActiveForm.Retrive val(.TextMatrix(.Row, .ColIndex("NoteID")))
        Else
                        If m_SearchType2 = 1 Then
                        
       
                 FrmBusinessJob.TxtPaymentVchrNo = (.TextMatrix(.Row, .ColIndex("NoteSerial")))
                          FrmBusinessJob.txtPaymentVchrValue = val(.TextMatrix(.Row, .ColIndex("NoteValue")))
                            
                        ElseIf SearchType = 7 Then
                        RSContract.TxtNotID.text = val(.TextMatrix(.Row, .ColIndex("NoteID")))
                       ' RSContract.TxtNotSreail1.text = val(.TextMatrix(.Row, .ColIndex("NoteSerial")))
                       ' RSContract.TxtNotVal.text = val(.TextMatrix(.Row, .ColIndex("NoteValue")))
                            ElseIf SearchType = 10 Then
                        FrmPayments2.TxtNotID.text = val(.TextMatrix(.Row, .ColIndex("NoteID")))
                     '  FrmPayments2.XPTxtVal.text = val(.TextMatrix(.Row, .ColIndex("NoteValue")))
                        FrmPayments2.TxtNotSreail1.text = val(.TextMatrix(.Row, .ColIndex("NoteSerial")))
                        FrmPayments2.TxtNotVal.text = val(.TextMatrix(.Row, .ColIndex("NoteValue")))
                        FrmPayments2.DcbUnitType.BoundText = val(.TextMatrix(.Row, .ColIndex("unittype")))
                        FrmPayments2.DcbUnitNo.BoundText = val(.TextMatrix(.Row, .ColIndex("UnitNo")))
                        FrmPayments2.DcbIqara.BoundText = val(.TextMatrix(.Row, .ColIndex("akarid")))
                        
                        ElseIf SearchType = 360 Or SearchType = 8008 Or SearchType = 3003 Or SearchType = 5005 Then
                           FrmExpenses40A.TxtVouSerial.text = (.TextMatrix(.Row, .ColIndex("NoteSerial")))
                        Else
                        mdifrmmain.ActiveForm.Retrive val(.TextMatrix(.Row, .ColIndex("NoteID")))
                        End If
        
        End If
    
    End With

End Sub

Private Sub Form_Activate()
If SearchType = 5 Or SearchType = 5005 Then
DCboCashType2.Visible = True
lbl(4).Visible = True
lbl(46).Visible = True
TxtOrder.Visible = True
Else
DCboCashType2.Visible = False
lbl(4).Visible = False
lbl(46).Visible = False
TxtOrder.Visible = False
End If
End Sub
Private Sub dcbAqarType_Change()
DcbUnitType_Change
End Sub
Private Sub DcbUnitType_Change()

Dim idd As Long
Dim idd1 As Long
   Set Dcombos = New ClsDataCombos

If val(dcbAqarType.BoundText) > 0 Then
idd = val(dcbAqarType.BoundText)

idd1 = val(DcbUnitType.BoundText)

Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"

End If

End Sub

Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
End Sub
Private Sub Form_Load()
    
    Dim GrdBack As New ClsBackGroundPic
  Dim StrSQL As String
YearMonth1
opt(1).value = True
    Set Dcombos = New ClsDataCombos
        Dcombos.GetIqar dcbAqarType
        Dcombos.getAkarUnit Me.DcbUnitType
       ' Dcombos.GetAccountingCodes Me.DcbAccount, , True
    DCboCashType2_Change
    CenterForm Me
    With DCboCashType2
    .Clear
    If SystemOptions.UserInterface = ArabicInterface Then
    .AddItem "≈·Ï ⁄„Ì·"
    .AddItem "≈·Ï „Ê—œ"
    .AddItem "„ﬁ«Ê· »«ÿ‰"
    .AddItem "„‘—Ê⁄"
    .AddItem "„ÊŸ›"
    .AddItem "Õ”«»"
    .AddItem "—Ê« »"
    .AddItem "„œ›Ê⁄«  „ﬁœ„…"
    .AddItem " „” Õﬁ«  ≈Ã«“…"
    .AddItem "”‰œ ’—› „ ⁄ÂœÌ‰"
    .AddItem "‰Â«Ì… Œœ„…"
    Else
     .AddItem "To Customer"
    .AddItem "To Vendor"
    .AddItem "sub-contractor"
    .AddItem "To Project"
    .AddItem "To Employee"
    .AddItem "To Acc."
    .AddItem "Salaries"
    .AddItem "Prepayments"
    .AddItem "Vacation Due"
    .AddItem "To Suppller."
    .AddItem "End Service"
    End If
    End With
    FormPostion Me, GetPostion
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DcboUsers
    Dcombos.GetCustomersSuppliers 0, Me.DcboCustomers
    Dcombos.GetExpensesType DcboExpensesType
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetRevenuesTypes Me.DcboRevenuesTypes
    Dcombos.GetCustomersSuppliers 57, Me.dcsupplier

               If SystemOptions.UserInterface = EnglishInterface Then
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng  FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                Else
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name  FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                End If
                  If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                   Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                 End If
                   StrSQL = StrSQL & GetAccountByBarnchUser
                   StrSQL = StrSQL & GetAccountCodeHiding
                fill_combo Me.DcbAccount, StrSQL
                
                

               If SystemOptions.UserInterface = EnglishInterface Then
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng  FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                Else
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name  FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "
                End If
                  If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                   Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                 End If
                 StrSQL = StrSQL + " And  ACCOUNTS.Account_Code In (Select Em.Account_Code from tblEmployee Em)   "
                fill_combo Me.DcbAccount2, StrSQL
                               
                
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

    With Me.CboPayMentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "«·ﬂ·"
        .ListIndex = 2
    End With

    With Me.FG
        Set .WallPaper = GrdBack.SearchWallpaper
        .AutoSize 0, .Cols - 1, False
    End With
    
    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DTPTo
    Me.CmdShowMoreOptions.value = False
    CmdShowMoreOptions_Click
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
Cmd_Click (1)
End Sub

Private Sub ChangeLang()
 Frame6.Caption = "Salaries"
 Label3.Caption = "Moth"
 Label4.Caption = "Year"
 CmdShowMoreOptions.Caption = "Adv.Search"
    With Me.CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheqye"
        .AddItem "All"
        .ListIndex = 2
    End With
 lbl(70).Caption = "End Service"
lbl(54).Caption = "Due Voucher"
lbl(46).Caption = "Order Exchane"
lbl(55).Caption = "Req .No"

    Me.Caption = "Search "
 lbl(4).Caption = "Type"
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
 XPLbl(4).Caption = "Type"
 lbl(3).Caption = "To"
 
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
        .TextMatrix(0, .ColIndex("Too")) = "Supp. Bill#"
         .TextMatrix(0, .ColIndex("Notes")) = "To"
         .TextMatrix(0, .ColIndex("ChqueNum")) = "Chque Num"
         
        
  
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

Private Sub DcbAccount_Click(Area As Integer)
TxtAccount.text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
End Sub

Private Sub Text3_Change()
 Dim Dcombos As New ClsDataCombos
    
    
 
 If Me.SearchType = 300 Then
    Dcombos.GetQuicSearch dcbAqarType, Text3, "FixedAssets", "Id", , , "FullCode"
Else
    Dcombos.GetQuicSearch dcbAqarType, Text3, "TblAqar", "Aqarid", "aqarname", "aqarname"
End If

    
End Sub

Private Sub TxtAccount_Change()
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.text)
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 0)
End Sub

Public Property Get SearchType() As Integer
    SearchType = m_SearchType
End Property

Public Property Let SearchType(ByVal vNewValue As Integer)
    m_SearchType = vNewValue
Frame1.Visible = False
dcbAqarType.Visible = False
DcbUnitType.Visible = False
DcbUnitNo.Visible = False
Label1(1).Visible = False
Label1(14).Visible = False
Label1(15).Visible = False
DcbAccount.Visible = False
Label1(0).Visible = False
TxtAccount.Visible = False
dcsupplier.Visible = False
Label1(2).Visible = False
    With Me.FG
    .ColHidden(.ColIndex("Account")) = True
  
        If m_SearchType = 3 Or m_SearchType = 333 Or m_SearchType = 360 Or m_SearchType = 6 Or m_SearchType = 2020 Then
        
            '3 «·»ÕÀ ⁄‰ «·„’—Ê›« 
            .ColHidden(.ColIndex("PaymentType")) = False
             .ColHidden(.ColIndex("CustName")) = True
             .ColHidden(.ColIndex("Too")) = False
             .ColHidden(.ColIndex("BoxName")) = False
             .ColHidden(.ColIndex("ManulaNO")) = False
            
            If m_SearchType = 333 Or m_SearchType = 360 Or m_SearchType = 6 Or m_SearchType = 2020 Then
            TxtAccount.Visible = True
            .ColHidden(.ColIndex("Account")) = False
             dcbAqarType.Visible = True
             DcbUnitType.Visible = True
            dcsupplier.Visible = True
            Label1(2).Visible = True
             DcbUnitNo.Visible = True
             Label1(1).Visible = True
             Label1(14).Visible = True
             Label1(15).Visible = True
             .ColHidden(.ColIndex("ManulaNO")) = True
             .ColHidden(.ColIndex("BoxName")) = True
             .ColHidden(.ColIndex("CustName")) = False
             .ColHidden(.ColIndex("Too")) = True
             DcbAccount.Visible = True
             Label1(0).Visible = True
            End If
            .ColHidden(.ColIndex("NoteCashingType")) = False
            .ColHidden(.ColIndex("BankName")) = True
            .ColHidden(.ColIndex("ChqueNum")) = True
            .ColHidden(.ColIndex("Notes")) = False
            Frame1.Visible = True
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(0, .ColIndex("Notes")) = "·«„—"
            Else
            .TextMatrix(0, .ColIndex("Notes")) = "To"
            End If
            .TextMatrix(0, .ColIndex("CustName")) = "«·⁄ﬁ«—"
            If m_SearchType = 333 Then
             If SystemOptions.UserInterface = ArabicInterface Then
                    Me.Caption = "«·»ÕÀ ⁄‰ «·„’—Ê›« "
            Else
                    Me.Caption = "Expense Search"
            End If
            End If
              If m_SearchType = 360 Then
             If SystemOptions.UserInterface = ArabicInterface Then
                    Me.Caption = "«·»ÕÀ ⁄‰  ›Ì… «·⁄Âœ"
            Else
                    Me.Caption = "Expense Search"
            End If
            End If
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
        ElseIf m_SearchType = 300 Then
            
            Me.Caption = "«·»ÕÀ ⁄‰ «· Œ·’ „‰ «·«’·"
            Label1(1).Caption = "«·«’·"
            Label1(1).Visible = True
            Text3.Visible = True
            dcbAqarType.Visible = True
            Dcombos.GetFixedAssets Me.dcbAqarType
            
        ElseIf m_SearchType = 30033 Then
            
            Me.Caption = "«·»ÕÀ ⁄‰ „‘ —Ì«  «’·"
            Label1(1).Caption = "«·«’·"
            Label1(1).Visible = True
            Text3.Visible = True
            dcbAqarType.Visible = True
            Dcombos.GetFixedAssets Me.dcbAqarType
        ElseIf m_SearchType = 4 Then
            '4 «·»ÕÀ ⁄‰ «·„ﬁ»Ê÷« 
            .ColHidden(.ColIndex("PaymentType")) = True
            .ColHidden(.ColIndex("CustName")) = False
                         If SystemOptions.UserInterface = ArabicInterface Then

            Me.Caption = "«·»ÕÀ ⁄‰ «·„ﬁ»Ê÷« "
            Else
            Me.Caption = "Cashing Account"
            End If
            Me.XPLbl(4).Visible = False
            Me.DcboExpensesType.Visible = False
        
            Me.XPLbl(5).Visible = True
            Me.DcboCustomers.Visible = True
            Me.CmdShowMoreOptions.Visible = True
            Me.CmdShowMoreOptions.value = False
        
            Me.DCboCashType.Visible = True
            Me.XPLbl(7).Visible = True
                          If SystemOptions.UserInterface = ArabicInterface Then
                              Me.XPLbl(7).Caption = "‰Ê⁄ «·„ﬁ»Ê÷« "
                           Else
                           Me.XPLbl(7).Caption = "Cashing Type "
                           End If
 If SystemOptions.UserInterface = ArabicInterface Then
            With Me.DCboCashType
                .Clear
                .AddItem "„‰ ⁄„Ì· «Ê „Ê—œ"
                .AddItem "„ﬁ«Ê· »«ÿ‰"
                .AddItem "≈Ì—œ«  ≈Œ—Ï"
                .AddItem "«·ﬂ·"
            End With
    Else
                With Me.DCboCashType
                .Clear
                .AddItem "Customer\supplier"
                .AddItem "Sub-contract"
                .AddItem "Other Revenue"
                .AddItem "all"
            End With
            
    
    End If
    
    

            DCboCashType.ListIndex = 3

            With Me.CboTrans
                .Clear
                .AddItem "›« Ê—… „»Ì⁄« "
                .AddItem "„— Ã⁄ „‘ —Ì« "
             
            End With

        ElseIf m_SearchType = 5 Or m_SearchType = 5005 Then
            '5 «·»ÕÀ «·„œ›Ê⁄« 
                Frame1.Visible = True
                 If SystemOptions.UserInterface = ArabicInterface Then

            Me.Caption = "«·»ÕÀ ⁄‰ «·„œ›Ê⁄« "
            Else
            Me.Caption = "Payment Search"
            End If
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
                  If SystemOptions.UserInterface = ArabicInterface Then
            Me.XPLbl(7).Caption = "‰Ê⁄ «·„œ›Ê⁄« "
Else
   Me.XPLbl(7).Caption = "Payment type "
End If
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
   StrSQL = "SELECT     dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value , dbo.Notes.Note_Value2, dbo.Notes.Remark,"
         StrSQL = StrSQL + "             dbo.Notes.NoteHijriDate, dbo.Notes.CashingType, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.ExpensesType.ID, dbo.ExpensesType.Name,"
         StrSQL = StrSQL + "             dbo.TblUsers.UserID, dbo.TblUsers.UserName, dbo.Transactions.Transaction_Serial, dbo.TransactionTypes.TransactionTypeName,"
         StrSQL = StrSQL + "             dbo.TblMaintenece.MaintananceID, dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.BanksData.BankName, dbo.Notes.ChqueNum, dbo.Notes.DueDate,"
         StrSQL = StrSQL + "             dbo.Notes.NoteCashingType, dbo.Notes.Water, dbo.Notes.Instrunce, dbo.Notes.comX, dbo.Notes.ComY, dbo.Notes.CommissionOut, dbo.Notes.rent,"
         StrSQL = StrSQL + "             dbo.Notes.commission, dbo.Notes.FIlterTotal, dbo.Notes.FilterID, dbo.Notes.StatusEarnest, dbo.Notes.Telephone, dbo.Notes.NoteId2, dbo.Notes.NoteSerial2,"
         StrSQL = StrSQL + "             dbo.Notes.unittype, dbo.Notes.UnitNo, dbo.Notes.akarid, dbo.Notes.[interval], dbo.Notes.intervaltype, dbo.Notes.renterName, dbo.Notes.AllowDateH,"
         StrSQL = StrSQL + "             dbo.Notes.AllowDate , dbo.Notes.ContNo, dbo.Notes.ContractNo, dbo.Notes.BookNo, dbo.Notes.ManualNO ,dbo.Notes.ManulaNO"
         StrSQL = StrSQL + "    FROM         dbo.Transactions RIGHT OUTER JOIN"
         StrSQL = StrSQL + "             dbo.TblCustemers RIGHT OUTER JOIN"
         StrSQL = StrSQL + "             dbo.BanksData RIGHT OUTER JOIN"
         StrSQL = StrSQL + "             dbo.TblBoxesData RIGHT OUTER JOIN"
         StrSQL = StrSQL + "             dbo.TblUsers INNER JOIN"
         StrSQL = StrSQL + "             dbo.ExpensesType RIGHT OUTER JOIN"
         StrSQL = StrSQL + "             dbo.Notes ON dbo.ExpensesType.ID = dbo.Notes.ExpensesID ON dbo.TblUsers.UserID = dbo.Notes.UserID ON dbo.TblBoxesData.BoxID = dbo.Notes.BoxID ON"
         StrSQL = StrSQL + "             dbo.BanksData.BankID = dbo.Notes.BankID ON dbo.TblCustemers.CusID = dbo.Notes.CusID ON"
         StrSQL = StrSQL + "             dbo.Transactions.Transaction_ID = dbo.Notes.Transaction_ID LEFT OUTER JOIN"
         StrSQL = StrSQL + "             dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type LEFT OUTER JOIN"
         StrSQL = StrSQL + "             dbo.TblMaintenece ON dbo.Notes.MaintananceID = dbo.TblMaintenece.MaintananceID"
               

       ' StrSQL = "SELECT dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial,dbo.Notes.NoteSerial1, dbo.Notes.Note_Value," & "dbo.Notes.Remark,dbo.Notes.NoteHijriDate, dbo.Notes.CashingType, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName," & "dbo.ExpensesType.ID,dbo.ExpensesType.Name, dbo.TblUsers.UserID, dbo.TblUsers.UserName, dbo.Transactions.Transaction_Serial," & "dbo.TransactionTypes.TransactionTypeName, dbo.TblMaintenece.MaintananceID, dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName," & "dbo.BanksData.BankName , dbo.Notes.ChqueNum, dbo.Notes.DueDate, dbo.Notes.NoteCashingType"
       ' StrSQL = StrSQL + " FROM dbo.Transactions RIGHT OUTER JOIN dbo.TblCustemers RIGHT OUTER JOIN dbo.BanksData RIGHT OUTER JOIN "
       ' StrSQL = StrSQL + " dbo.TblBoxesData RIGHT OUTER JOIN dbo.TblUsers INNER JOIN dbo.ExpensesType RIGHT OUTER JOIN "
       ' StrSQL = StrSQL + " dbo.Notes ON dbo.ExpensesType.ID = dbo.Notes.ExpensesID ON dbo.TblUsers.UserID = dbo.Notes.UserID ON "
       ' StrSQL = StrSQL + " dbo.TblBoxesData.BoxID = dbo.Notes.BoxID ON dbo.BanksData.BankID = dbo.Notes.BankID ON dbo.TblCustemers.CusID = dbo.Notes.CusID ON"
       ' StrSQL = StrSQL + " dbo.Transactions.Transaction_ID = dbo.Notes.Transaction_ID LEFT OUTER JOIN "
       ' StrSQL = StrSQL + " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type LEFT OUTER JOIN"
       ' StrSQL = StrSQL + " dbo.TblMaintenece ON dbo.Notes.MaintananceID = dbo.TblMaintenece.MaintananceID "
    
        'Debug.Print Replace(StrSQL, "dbo.", "", , , vbTextCompare)
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial, Notes.Note_Value," & "Notes.Remark,Notes.NoteHijriDate, Notes.CashingType, TblCustemers.CusID, TblCustemers.CusName," & "ExpensesType.ID,ExpensesType.Name, TblUsers.UserID, TblUsers.UserName, Transactions.Transaction_Serial," & "TransactionTypes.TransactionTypeName, TblMaintenece.MaintananceID, TblBoxesData.BoxID, TblBoxesData.BoxName," & "BanksData.BankName , Notes.ChqueNum, Notes.DueDate, Notes.NoteCashingType "
        StrSQL = StrSQL + " FROM TransactionTypes RIGHT JOIN (Transactions RIGHT JOIN (TblMaintenece RIGHT JOIN " & "(TblCustemers RIGHT JOIN (TblBoxesData RIGHT JOIN (ExpensesType RIGHT JOIN (BanksData RIGHT JOIN " & "(TblUsers INNER JOIN Notes ON TblUsers.UserID = Notes.UserID) ON BanksData.BankID = Notes.BankID) ON " & "ExpensesType.ID = Notes.ExpensesID) ON TblBoxesData.BoxID = Notes.BoxID) ON TblCustemers.CusID = Notes.CusID) " & "ON TblMaintenece.MaintananceID = Notes.MaintananceID) ON Transactions.Transaction_ID = Notes.Transaction_ID) " & "ON TransactionTypes.Transaction_Type = Transactions.Transaction_Type "
    End If

    IntNoteType = Me.SearchType
    If IntNoteType = 3 Or IntNoteType = 80 Or IntNoteType = 8008 Or IntNoteType = 3003 Or IntNoteType = 300 Or IntNoteType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
 
        StrSQL = "SELECT   dbo.notes_all.bill_Type,dbo.notes_all.A_NoteID ,   dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial, dbo.notes_all.NoteSerial1, dbo.notes_all.Note_Value, "
        StrSQL = StrSQL + " dbo.notes_all.Remark, dbo.notes_all.NoteHijriDate, dbo.notes_all.CashingType, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.ExpensesType.ID,"
        StrSQL = StrSQL + " dbo.ExpensesType.Name, dbo.TblUsers.UserID, dbo.TblUsers.UserName, dbo.Transactions.Transaction_Serial, dbo.TransactionTypes.TransactionTypeName,"
        StrSQL = StrSQL + " dbo.TblMaintenece.MaintananceID, dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.BanksData.BankName, dbo.notes_all.ChqueNum,"
        StrSQL = StrSQL + " dbo.notes_all.DueDate , dbo.notes_all.NoteCashingType, dbo.notes_all.too "
        ', dbo.Notes.TxtEndService, dbo.Notes.OrderIDD,  dbo.Notes.Due, dbo.Notes.TxtOrderSuppler"
        StrSQL = StrSQL + " FROM         dbo.Transactions RIGHT OUTER JOIN"
        StrSQL = StrSQL + " dbo.TblCustemers RIGHT OUTER JOIN"
        StrSQL = StrSQL + "  dbo.BanksData RIGHT OUTER JOIN"
        StrSQL = StrSQL + " dbo.TblBoxesData RIGHT OUTER JOIN"
        StrSQL = StrSQL + "  dbo.TblUsers INNER JOIN"
        StrSQL = StrSQL + "  dbo.ExpensesType RIGHT OUTER JOIN"
        StrSQL = StrSQL + "  dbo.notes_all ON dbo.ExpensesType.ID = dbo.notes_all.ExpensesID ON dbo.TblUsers.UserID = dbo.notes_all.UserID ON"
        StrSQL = StrSQL + "  dbo.TblBoxesData.BoxID = dbo.notes_all.BoxID ON dbo.BanksData.BankID = dbo.notes_all.BankID ON dbo.TblCustemers.CusID = dbo.notes_all.CusID ON"
        StrSQL = StrSQL + "  dbo.Transactions.Transaction_ID = dbo.notes_all.Transaction_ID LEFT OUTER JOIN"
        StrSQL = StrSQL + " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type LEFT OUTER JOIN"
        StrSQL = StrSQL + "   dbo.TblMaintenece ON dbo.notes_all.MaintananceID = dbo.TblMaintenece.MaintananceID  "

    End If
    If IntNoteType = 333 Then
     StrSQL = " SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial, dbo.notes_all.Note_Value, dbo.notes_all.BankID, "
     StrSQL = StrSQL & "                  dbo.notes_all.ChqueNum, dbo.notes_all.DueDate, dbo.notes_all.UserID, dbo.notes_all.Remark, dbo.notes_all.ExpensesID, dbo.notes_all.BoxID, dbo.notes_all.too,"
     StrSQL = StrSQL & "                   dbo.notes_all.note_value_by_characters, dbo.notes_all.general_des, dbo.notes_all.NoteSerial1, dbo.notes_all.ToPriodDateH, dbo.notes_all.FrmPriodDateH,"
     StrSQL = StrSQL & "                   dbo.notes_all.ToPriodDate, dbo.notes_all.FrmPriodDate, dbo.notes_all.Iqar, dbo.notes_all.UnitType, dbo.notes_all.NoteDateH, dbo.notes_all.CashingType,"
     StrSQL = StrSQL & "                   dbo.notes_all.NoteCashingType, dbo.notes_all.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblExpensesDet.Unitss,"
     StrSQL = StrSQL & "                   dbo.TblExpensesDet.StrUnit, dbo.TblExpensesDet.AccountCode, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng,"
     StrSQL = StrSQL & "                   dbo.TblExpensesDet.des, dbo.TblExpensesDet.order_no, dbo.TblExpensesDet.opr_fullcode, dbo.TblExpensesDet.[value], dbo.TblExpensesDet.iqarid,"
     StrSQL = StrSQL & "                   dbo.TblAqar.aqarname, dbo.TblExpensesDet.uintid, dbo.TblAqarDetai.unitno, dbo.TblExpensesDet.type, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee,"
     StrSQL = StrSQL & "                   dbo.notes_all.NoteHijriDate , dbo.BanksData.BankName, dbo.BanksData.BankNamee, dbo.TblUsers.UserName"
     StrSQL = StrSQL & "          FROM         dbo.TblUsers RIGHT OUTER JOIN"
     StrSQL = StrSQL & "                   dbo.notes_all ON dbo.TblUsers.UserID = dbo.notes_all.UserID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                   dbo.BanksData ON dbo.notes_all.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                   dbo.TblAkarUnit RIGHT OUTER JOIN"
     StrSQL = StrSQL & "                   dbo.TblExpensesDet ON dbo.TblAkarUnit.id = dbo.TblExpensesDet.type LEFT OUTER JOIN"
     StrSQL = StrSQL & "                   dbo.TblAqarDetai ON dbo.TblExpensesDet.uintid = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
     StrSQL = StrSQL & "                   dbo.TblAqar ON dbo.TblExpensesDet.iqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
     StrSQL = StrSQL & "                   dbo.ACCOUNTS ON dbo.TblExpensesDet.AccountCode = dbo.ACCOUNTS.Account_Code ON dbo.notes_all.NoteID = dbo.TblExpensesDet.ExpID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                   dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
     
    End If
      If IntNoteType = 360 Then
     StrSQL = " SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial, dbo.notes_all.Note_Value, dbo.notes_all.BankID, "
     StrSQL = StrSQL & "                 dbo.notes_all.ChqueNum, dbo.notes_all.DueDate, dbo.notes_all.UserID, dbo.notes_all.Remark, dbo.notes_all.ExpensesID, dbo.notes_all.BoxID, dbo.notes_all.too,"
     StrSQL = StrSQL & "                 dbo.notes_all.note_value_by_characters, dbo.notes_all.general_des, dbo.notes_all.NoteSerial1, dbo.notes_all.ToPriodDateH, dbo.notes_all.FrmPriodDateH,"
     StrSQL = StrSQL & "                 dbo.notes_all.ToPriodDate, dbo.notes_all.FrmPriodDate, dbo.notes_all.Iqar, dbo.notes_all.UnitType, dbo.notes_all.NoteDateH, dbo.notes_all.CashingType,"
     StrSQL = StrSQL & "                 dbo.notes_all.NoteCashingType, dbo.notes_all.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.ACCOUNTS.Account_Name,"
     StrSQL = StrSQL & "                 dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.TblAqarDetai.unitno, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee,"
     StrSQL = StrSQL & "                 dbo.notes_all.NoteHijriDate, dbo.BanksData.BankName, dbo.BanksData.BankNamee, dbo.TblUsers.UserName, dbo.TblExpensesDet301.UnitType AS Expr1,"
     StrSQL = StrSQL & "                 dbo.TblExpensesDet301.AccountCode, dbo.TblExpensesDet301.UnitNo AS Expr2, dbo.TblExpensesDet301.[Value], dbo.TblExpensesDet301.Unitss,"
     StrSQL = StrSQL & "                 dbo.TblExpensesDet301.Aqarid , dbo.TblAqar.aqarname"
     StrSQL = StrSQL & "        FROM         dbo.BanksData RIGHT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblAkarUnit RIGHT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.ACCOUNTS RIGHT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblAqar RIGHT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblExpensesDet301 ON dbo.TblAqar.Aqarid = dbo.TblExpensesDet301.Aqarid RIGHT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.notes_all ON dbo.TblExpensesDet301.ExpID = dbo.notes_all.NoteID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblAqarDetai ON dbo.TblExpensesDet301.UnitNo = dbo.TblAqarDetai.Id ON dbo.ACCOUNTS.Account_Code = dbo.TblExpensesDet301.AccountCode ON"
     StrSQL = StrSQL & "                 dbo.TblAkarUnit.id = dbo.TblExpensesDet301.UnitType LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblUsers ON dbo.notes_all.UserID = dbo.TblUsers.UserID ON dbo.BanksData.BankID = dbo.notes_all.BankID LEFT OUTER JOIN"
     StrSQL = StrSQL & "                 dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
    End If
    
  Dim tep As Integer
   If IntNoteType = 333 Then IntNoteType = 3
If IntNoteType = 3003 Then IntNoteType = 3
If IntNoteType = 8008 Then IntNoteType = 80
If IntNoteType = 300 Then IntNoteType = 8028
If IntNoteType = 30033 Then IntNoteType = 80

If IntNoteType = 360 Then IntNoteType = 350
    If IntNoteType = 333 Or IntNoteType = 3 Or IntNoteType = 80 Or IntNoteType = 350 Or IntNoteType = 360 Or IntNoteType = 350 Or IntNoteType = 8028 Or IntNoteType = 80 Then
        StrWhere = " Where notes_all.NoteType=" & IntNoteType & ""
    ElseIf IntNoteType = 300 Then
        StrWhere = " Where notes_all.NoteType=80"
        
    ElseIf IntNoteType = 8063 Then
   StrWhere = " Where notes_all.NoteType=85" ' = 8063
    
    Else
    If IntNoteType = 7 Then
     'StrWhere = StrWhere + " AND  (CashingType=9 )"
   ' StrWhere = " Where CashingType=9"
    End If
    
    
   
    tep = IntNoteType
    If IntNoteType = 6 Then IntNoteType = 4
    If IntNoteType = 5005 Then IntNoteType = 5
     If IntNoteType = 7 Then IntNoteType = 4
     If IntNoteType = 10 Then IntNoteType = 4
      If IntNoteType = 2020 Then IntNoteType = 5
        StrWhere = " Where  Notes.NoteType=" & IntNoteType & ""
       
        If tep = 6 Then
       ' StrWhere = StrWhere + " AND    dbo.Notes.CashingType >7"
        End If
        
        If val(DCboCashType2.ListIndex) <> -1 Then
        StrWhere = StrWhere + " AND    dbo.Notes.CashingType =" & val(DCboCashType2.ListIndex) & ""
        End If
        
       If SearchType = 2020 Then
            If dcbAqarType.text <> "" And val(dcbAqarType.BoundText) <> 0 Then
                    StrWhere = StrWhere + " AND dbo.notes.akarid = " & val(dcbAqarType.BoundText) & ""
            End If
        If DcbUnitType.text <> "" And val(DcbUnitType.BoundText) <> 0 Then
            StrWhere = StrWhere + " AND dbo.notes.unittype = " & val(DcbUnitType.BoundText) & ""
        End If
        If DcbUnitNo.text <> "" Then
            StrWhere = StrWhere + " AND dbo.notes.unitno =" & val(DcbUnitNo.BoundText)
        End If
       
        If dcsupplier.text <> "" And val(dcsupplier.BoundText) <> 0 Then
           StrWhere = StrWhere + " AND dbo.notes.akarid In  (Select  Aqarid from TblAqar Where ownerid = " & val(dcsupplier.BoundText) & ")"
        End If
        
     ' End If
       
       
    End If
              
        If val(DCboCashType2.ListIndex) = 6 Then
        If val(CmbMonth1.ListIndex) <> -1 Then
        StrWhere = StrWhere + " AND    dbo.Notes.PayrollMonth =" & val(CmbMonth1.ListIndex) + 1 & ""
        End If
        If val(CboYear1.ListIndex) <> -1 Then
        StrWhere = StrWhere + " AND    dbo.Notes.PayrollYear =" & val(CboYear1.ListIndex) & ""
        End If
        End If
        If val(TxtOrder.text) <> 0 Then
        StrWhere = StrWhere + " AND    dbo.Notes.OrderIDD=" & TxtOrder.text & ""
        End If
        If IntNoteType = 5 Then
        If val(DCboCashType2.ListIndex) = 8 Then
          If val(TxtDue.text) <> 0 Then
          StrWhere = StrWhere + " AND    dbo.Notes.Due=" & TxtDue.text & ""
         End If
         
         ElseIf val(DCboCashType2.ListIndex) = 9 Then
         
          If val(TxtOrderSuppler.text) <> 0 Then
          StrWhere = StrWhere + " AND    dbo.Notes.TxtOrderSuppler=" & TxtOrderSuppler.text & ""
         End If
         ElseIf val(DCboCashType2.ListIndex) = 10 Then
         
         If val(TxtEndService.text) <> 0 Then
          StrWhere = StrWhere + " AND    dbo.Notes.TxtEndService=" & TxtEndService.text & ""
         End If
        End If
        End If
    End If
If tep = 7 Or tep = 10 Then
StrWhere = StrWhere + " AND  (CashingType=9 )"

End If
    If IntNoteType = 80 Then
       ' StrWhere = StrWhere + " AND   notes_all.bill_Type<>2"
    ElseIf IntNoteType = 300 Then
        StrWhere = StrWhere + " AND   notes_all.bill_Type=2"
    
    End If

    If IntNoteType = 3 Or IntNoteType = 333 Or Me.SearchType = 360 Then

        '«·„’—Ê›« 
        If Me.DcboExpensesType.BoundText <> "" Then
            StrWhere = StrWhere + " AND  Notes.ExpensesID=" & Me.DcboExpensesType.BoundText & ""
        End If
  If SearchType = 333 Then
     If dcbAqarType.text <> "" And val(dcbAqarType.BoundText) <> 0 Then
     StrWhere = StrWhere + " AND dbo.TblExpensesDet.iqarid = " & val(dcbAqarType.BoundText) & ""
   End If
   If DcbUnitType.text <> "" And val(DcbUnitType.BoundText) <> 0 Then
     StrWhere = StrWhere + " AND dbo.TblExpensesDet.type = " & val(DcbUnitType.BoundText) & ""
   End If
  If DcbUnitNo.text <> "" Then
     StrWhere = StrWhere + " AND dbo.TblExpensesDet.Unitss LIKE N'%" & DcbUnitNo.text & "%'"
   End If
     If DcbAccount.text <> "" Then
     StrWhere = StrWhere + " AND dbo.TblExpensesDet.AccountCode = '" & Me.DcbAccount.BoundText & "'"
   End If
   
  End If
    If SearchType = 360 Then
     If dcbAqarType.text <> "" And val(dcbAqarType.BoundText) <> 0 Then
     StrWhere = StrWhere + " AND dbo.TblExpensesDet301.Aqarid = " & val(dcbAqarType.BoundText) & ""
   End If
   If DcbUnitType.text <> "" And val(DcbUnitType.BoundText) <> 0 Then
     StrWhere = StrWhere + " AND dbo.TblExpensesDet301.UnitType = " & val(DcbUnitType.BoundText) & ""
   End If
  If DcbUnitNo.text <> "" Then
     StrWhere = StrWhere + " AND dbo.TblExpensesDet301.Unitss LIKE N'%" & DcbUnitNo.text & "%'"
   End If
     If DcbAccount.text <> "" Then
     StrWhere = StrWhere + " AND dbo.TblExpensesDet301.AccountCode = '" & Me.DcbAccount.BoundText & "'"
   End If
   
  End If
    ElseIf IntNoteType = 4 Then




      If SearchType = 6 Or SearchType = 2020 Then
     If dcbAqarType.text <> "" And val(dcbAqarType.BoundText) <> 0 Then
     StrWhere = StrWhere + " AND dbo.notes.akarid = " & val(dcbAqarType.BoundText) & ""
   End If
   If DcbUnitType.text <> "" And val(DcbUnitType.BoundText) <> 0 Then
     StrWhere = StrWhere + " AND dbo.notes.unittype = " & val(DcbUnitType.BoundText) & ""
   End If
  If DcbUnitNo.text <> "" Then
     StrWhere = StrWhere + " AND dbo.notes.unitno =" & val(DcbUnitNo.BoundText)
   End If
    
     If dcsupplier.text <> "" And val(dcsupplier.BoundText) <> 0 Then
        StrWhere = StrWhere + " AND dbo.notes.akarid In  (Select  Aqarid from TblAqar Where ownerid = " & val(dcsupplier.BoundText) & ")"
     End If
     
  ' End If
    
    
 End If
   
   
        '«·„ﬁ»Ê÷« 
        If Me.DCboCashType.ListIndex = 0 Then
            StrWhere = StrWhere + " AND  (CashingType=0 OR  CashingType=1)"
           
        ElseIf Me.DCboCashType.ListIndex = 1 Then
            StrWhere = StrWhere + " AND  (CashingType=2)"
        ElseIf Me.DCboCashType.ListIndex = 2 Then
            '«·≈Ì—«œ«  «·√Œ—Ï
            StrWhere = StrWhere + " AND  (CashingType=3)"
        End If
        If TxtManulaNO.text <> "" Then
           StrWhere = StrWhere + " AND (dbo.Notes.ManulaNO Like '%" & Trim(Me.TxtManulaNO.text) & "%')"
        End If
    ElseIf IntNoteType = 5 Or IntNoteType = 50 Then
  
        '«·„œ›Ê⁄« 
        If Me.DCboCashType.ListIndex = 0 Then
            StrWhere = StrWhere + " AND  (CashingType=0 OR  CashingType=1)"
        ElseIf Me.DCboCashType.ListIndex = 1 Then
            StrWhere = StrWhere + " AND  (CashingType=2)"
        End If
        
        
        If m_SearchType2 = 1 Then
          StrWhere = StrWhere + " AND  (person='" & person & "')"
        End If
        
    End If

    If Trim(Me.TxtSerial.text) <> "" Then
        If Me.SearchType = 3333 Or Me.SearchType = 3 Or Me.SearchType = 333 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Or Me.SearchType = 360 Then
            StrWhere = StrWhere + " AND dbo.notes_all.NoteSerial1=" & val(Me.TxtSerial.text) & ""
        Else
            StrWhere = StrWhere + " AND dbo.Notes.NoteSerial1=" & val(Me.TxtSerial.text) & ""
        End If
    End If

    If Me.DcboBox.BoundText <> "" Then
        If Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 360 Or Me.SearchType = 8063 Then
            StrWhere = StrWhere + " AND notes_all.BoxID =" & Me.DcboBox.BoundText & ""
        Else
            StrWhere = StrWhere + " AND Notes.BoxID =" & Me.DcboBox.BoundText & ""
        End If
    End If
    
    If Me.SearchType = 300 Then
        If dcbAqarType.text <> "" Then
            StrWhere = StrWhere + " AND notes_all.FAID =" & val(Me.dcbAqarType.BoundText) & ""
        End If
    End If
    
    If val(Me.TxtValue.text) > 0 Then
        If Me.opt(1).value = True Then
            If Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
                StrWhere = StrWhere + " AND notes_all.Note_Value =" & val(Me.TxtValue.text) & ""
           
            ElseIf Me.SearchType <> 333 And Me.SearchType <> 360 Then
               If Me.SearchType = 6 Or Me.SearchType = 7 Then
                StrWhere = StrWhere + " AND Notes.Note_Value2 =" & val(Me.TxtValue.text) & ""
              Else
                StrWhere = StrWhere + " AND Notes.Note_Value =" & val(Me.TxtValue.text) & ""
              End If
             
              
            ElseIf Me.SearchType = 333 Then
                StrWhere = StrWhere + " AND dbo.TblExpensesDet.value =" & val(Me.TxtValue.text) & ""
           ElseIf Me.SearchType = 360 Then
                StrWhere = StrWhere + " AND dbo.TblExpensesDet301.value =" & val(Me.TxtValue.text) & ""
             End If

        ElseIf Me.opt(0).value = True Then

            If Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
                StrWhere = StrWhere + " AND notes_all.Note_Value >" & val(Me.TxtValue.text) & ""
               
              
           ElseIf Me.SearchType <> 333 And Me.SearchType <> 360 Then
                
                             If Me.SearchType = 6 Or Me.SearchType = 7 Then
                StrWhere = StrWhere + " AND Notes.Note_Value2>" & val(Me.TxtValue.text) & ""
              Else
                StrWhere = StrWhere + " AND Notes.Note_Value >" & val(Me.TxtValue.text) & ""
              End If
 
 
               ' StrWhere = StrWhere + " AND Notes.Note_Value >" & val(Me.TxtValue.Text) & ""
                
                
           ElseIf Me.SearchType = 333 Then
                StrWhere = StrWhere + " AND dbo.TblExpensesDet.value >" & val(Me.TxtValue.text) & ""
           ElseIf Me.SearchType = 360 Then
                StrWhere = StrWhere + " AND dbo.TblExpensesDet301.value >" & val(Me.TxtValue.text) & ""
           End If
        Else

            If Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
                StrWhere = StrWhere + " AND notes_all.Note_Value <" & val(Me.TxtValue.text) & ""
            ElseIf Me.SearchType <> 333 And Me.SearchType <> 360 Then
                      If Me.SearchType = 6 Or Me.SearchType = 7 Then
                StrWhere = StrWhere + " AND Notes.Note_Value2 <" & val(Me.TxtValue.text) & ""
              Else
                StrWhere = StrWhere + " AND Notes.Note_Value <" & val(Me.TxtValue.text) & ""
              End If


                StrWhere = StrWhere + " AND Notes.Note_Value <" & val(Me.TxtValue.text) & ""
            ElseIf Me.SearchType = 333 Then
                StrWhere = StrWhere + " AND dbo.TblExpensesDet.value <" & val(Me.TxtValue.text) & ""
            ElseIf Me.SearchType = 360 Then
                StrWhere = StrWhere + " AND dbo.TblExpensesDet301.value <" & val(Me.TxtValue.text) & ""
             End If
        End If
    End If

    If Me.DcboUsers.BoundText <> "" Then
        If Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Or Me.SearchType = 360 Then
            StrWhere = StrWhere + " AND notes_all.UserID=" & Me.DcboUsers.BoundText & ""
        Else
            StrWhere = StrWhere + " AND Notes.UserID=" & Me.DcboUsers.BoundText & ""
        End If

    End If

    If txtto.text <> "" Then
        If Me.SearchType = 360 Or Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
            StrWhere = StrWhere + " AND notes_all.too='" & txtto.text & "'"
        End If

    End If

    If DcboCustomers.Visible = True Then
        If Me.DcboCustomers.BoundText <> "" Then
            If Me.SearchType = 360 Or Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
                StrWhere = StrWhere + " AND notes_all.CusID=" & Me.DcboCustomers.BoundText & ""
            Else
                StrWhere = StrWhere + " AND Notes.CusID=" & Me.DcboCustomers.BoundText & ""
            End If

        End If
    End If

    If Not IsNull(Me.DTPFrom.value) Then
        If Me.SearchType = 360 Or Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
            StrWhere = StrWhere + " AND  notes_all.NoteDate >=" & SQLDate(Me.DTPFrom.value, True) & ""
        Else
            StrWhere = StrWhere + " AND  Notes.NoteDate >=" & SQLDate(Me.DTPFrom.value, True) & ""
        
        End If
    End If

    If Not IsNull(Me.DTPTo.value) Then
        If Me.SearchType = 360 Or Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
            StrWhere = StrWhere + " AND  notes_all.NoteDate <=" & SQLDate(Me.DTPTo.value, True) & ""
        Else
            StrWhere = StrWhere + " AND  Notes.NoteDate <=" & SQLDate(Me.DTPTo.value, True) & ""
        End If
    End If

    If Me.ChkTrans.value = vbChecked Then
        If Me.SearchType = 360 Or Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
            StrWhere = StrWhere + " AND(notes_all.Transaction_ID is not null Or  Notes.MaintananceID is not null)"
        Else
            StrWhere = StrWhere + " AND(Notes.Transaction_ID is not null Or  Notes.MaintananceID is not null)"
        End If

    End If

    '----------------------
    'More Options Part
    If CboPayMentType.ListIndex = 0 Then
        If Me.SearchType = 360 Or Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
            StrWhere = StrWhere + " AND  (notes_all.NoteCashingType=0)"
        Else
            StrWhere = StrWhere + " AND  (NOTES.NoteCashingType=0)"
        End If

    ElseIf Me.CboPayMentType.ListIndex = 1 Then

        If Me.SearchType = 360 Or Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
            StrWhere = StrWhere + " AND  (notes_all.NoteCashingType=1)"
        Else
            StrWhere = StrWhere + " AND  (NOTES.NoteCashingType=1)"
        End If
    End If

    If Me.DcboBankName.BoundText <> "" Then
        If Me.SearchType = 360 Or Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
            StrWhere = StrWhere + " AND  (notes_all.BankID=" & val(Me.DcboBankName.BoundText) & ")"
        Else
            StrWhere = StrWhere + " AND  (NOTES.BankID=" & val(Me.DcboBankName.BoundText) & ")"
        End If
    End If

    If Trim$(Me.TxtChequeNumber.text) <> "" Then
        If Me.SearchType = 360 Or Me.SearchType = 333 Or Me.SearchType = 3 Or Me.SearchType = 8008 Or Me.SearchType = 3003 Or Me.SearchType = 80 Or Me.SearchType = 300 Or Me.SearchType = 30033 Or Me.SearchType = 350 Or Me.SearchType = 8063 Then
            StrWhere = StrWhere + " AND  (notes_all.ChqueNum='" & Trim$(Me.TxtChequeNumber.text) & "')"
        Else
            StrWhere = StrWhere + " AND  (NOTES.ChqueNum='" & Trim$(Me.TxtChequeNumber.text) & "')"
        End If
    End If
If Me.SearchType = 5 Then
  If Trim(DcbAccount2.text) <> "" Then
    StrWhere = StrWhere + " AND  (NOTES.EmpAccountCode='" & Trim(Me.DcbAccount2.BoundText) & "')"
  End If
  End If
    If Me.CboTrans.ListIndex <> -1 Then
        If Me.SearchType = 4 Then '„ﬁ»Ê÷« 
            If Me.CboTrans.ListIndex = 0 Then '"›« Ê—… „»Ì⁄« "
                StrWhere = StrWhere + " AND (Transactions.Transaction_Type=2)"
            ElseIf Me.CboTrans.ListIndex = 1 Then '"„— Ã⁄ „‘ —Ì« "
                StrWhere = StrWhere + " AND (Transactions.Transaction_Type=5)"
            ElseIf Me.CboTrans.ListIndex = 2 Then '"’Ì«‰…"
            ElseIf Me.CboTrans.ListIndex = 3 Then '"Œœ„« "
            End If

        ElseIf Me.SearchType = 5 Or SearchType = 5005 Then

            If Me.CboTrans.ListIndex = 0 Then '"›« Ê—… „‘ —Ì« "
                StrWhere = StrWhere + " AND (Transactions.Transaction_Type=1)"
            ElseIf Me.CboTrans.ListIndex = 1 Then '"„— Ã⁄ „»Ì⁄« "
                StrWhere = StrWhere + " AND (Transactions.Transaction_Type=9)"
            End If
        End If

        If Trim(Me.TxtTransSerial.text) <> "" Then
            If Me.chk.value = vbChecked Then
                StrWhere = StrWhere + " AND (Transactions.Transaction_Serial='" & Trim(Me.TxtTransSerial.text) & "')"
            ElseIf Me.chk.value = vbUnchecked Then
                StrWhere = StrWhere + " AND (Transactions.Transaction_Serial Like '%" & Trim(Me.TxtTransSerial.text) & "%')"
            End If
        End If
    
    End If
    If Me.SearchType = 3 Or Me.SearchType = 333 Then
     If Trim(Me.txtremark.text) <> "" Then
     StrWhere = StrWhere + " AND (notes_all.Remark  Like '%" & Trim(Me.txtremark.text) & "%')"
    End If
  End If



    If Me.SearchType = 5 Or SearchType = 5005 Then
     If Trim(Me.txtremark.text) <> "" Then
     StrWhere = StrWhere + " AND (Notes.person  Like '%" & Trim(Me.txtremark.text) & "%')"
    End If
  End If
  
  If SearchType = 7 Or SearchType = 10 Then
   StrWhere = StrWhere + " AND (Notes.PayedOrBon IS NULL)"
  End If
    If SearchType = 7 Then
   StrWhere = StrWhere + " AND (Notes.StatusEarnest =0)"
  End If
      If Me.SearchType = 2020 Then
    
        StrWhere = StrWhere & "   and  NoteType=5 and cashingtype<=8 "
        StrWhere = StrWhere & "     AND Notes.branch_no in(" & Current_branchSql & ")"
        StrWhere = StrWhere & " and  not (  (akarid is null )  and   (IqarID2 is null )  and   (NoteOrBonID is null ) )  "
    End If
    If Me.SearchType = 360 Or Me.SearchType = 333 Or IntNoteType = 3 Or IntNoteType = 80 Or IntNoteType = 300 Or Me.SearchType = 30033 Or IntNoteType = 350 Or IntNoteType = 8063 Or IntNoteType = 8028 Or IntNoteType = 80 Then
        StrSQL = StrSQL + StrWhere + " Order By dbo.notes_all.NoteSerial1"
    Else
        StrSQL = StrSQL + StrWhere + " Order By dbo.Notes.NoteSerial1"
    End If

    Build_Sql = StrSQL
End Function

