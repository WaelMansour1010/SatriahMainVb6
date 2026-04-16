VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmPayments 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "”‰œ ’—› - «·„œ›Ê⁄«   "
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14730
   HelpContextID   =   390
   Icon            =   "FrmPayments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   14730
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
   Begin VB.TextBox txtAcceptianPeriod 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2970
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   451
      Top             =   1950
      Width           =   855
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H0080FFFF&
      Caption         =   "»Ì«‰«  ›Ê« Ì— «·„‘ —Ì« "
      Height          =   6975
      Index           =   0
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   403
      Top             =   1050
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CommandButton Command10 
         BackColor       =   &H8000000B&
         Caption         =   "«·€«¡ «·”œ«œ"
         Height          =   315
         Index           =   0
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   405
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   10800
         RightToLeft     =   -1  'True
         TabIndex        =   404
         Top             =   300
         Width           =   1200
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   5460
         Left            =   120
         TabIndex        =   406
         Top             =   630
         Width           =   12360
         _cx             =   21802
         _cy             =   9631
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
         Cols            =   23
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPayments.frx":038A
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
         Left            =   12360
         RightToLeft     =   -1  'True
         TabIndex        =   409
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   408
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6240
         Width           =   8775
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «·›Ê« Ì—"
         Height          =   255
         Index           =   0
         Left            =   9600
         RightToLeft     =   -1  'True
         TabIndex        =   407
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H0080FFFF&
      Caption         =   "»Ì«‰«  «·›Ê« Ì— «·„«·Ì…"
      Height          =   6975
      Left            =   1230
      RightToLeft     =   -1  'True
      TabIndex        =   410
      Top             =   1050
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CheckBox Check18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   10800
         RightToLeft     =   -1  'True
         TabIndex        =   412
         Top             =   300
         Width           =   1200
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid1 
         Height          =   5460
         Left            =   120
         TabIndex        =   413
         Top             =   600
         Width           =   12360
         _cx             =   21802
         _cy             =   9631
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
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPayments.frx":0751
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
      Begin VB.CommandButton Command6 
         BackColor       =   &H8000000B&
         Caption         =   "«·€«¡ «·”œ«œ"
         Height          =   315
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   411
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «·›Ê« Ì—"
         Height          =   255
         Left            =   9600
         RightToLeft     =   -1  'True
         TabIndex        =   416
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   415
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6240
         Width           =   8775
      End
      Begin VB.Label Label17 
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
         Left            =   12360
         RightToLeft     =   -1  'True
         TabIndex        =   414
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H0080FFFF&
      Caption         =   "»Ì«‰«  «·„” Œ·’« "
      Height          =   6975
      Index           =   1
      Left            =   1230
      RightToLeft     =   -1  'True
      TabIndex        =   396
      Top             =   1050
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   10800
         RightToLeft     =   -1  'True
         TabIndex        =   398
         Top             =   300
         Width           =   1200
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H8000000B&
         Caption         =   "«·€«¡ «·”œ«œ"
         Height          =   315
         Index           =   1
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   397
         Top             =   240
         Width           =   1695
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
         Height          =   5700
         Left            =   0
         TabIndex        =   399
         Top             =   720
         Width           =   12360
         _cx             =   21802
         _cy             =   10054
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
         Cols            =   19
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPayments.frx":09F2
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
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «·›Ê« Ì—"
         Height          =   255
         Index           =   2
         Left            =   9600
         RightToLeft     =   -1  'True
         TabIndex        =   402
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label28 
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
         Left            =   12360
         RightToLeft     =   -1  'True
         TabIndex        =   400
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   401
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6480
         Width           =   8775
      End
   End
   Begin VB.TextBox TxtVATValue 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   447
      Top             =   4710
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtTotalWithVat 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5220
      RightToLeft     =   -1  'True
      TabIndex        =   445
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox TxtVAt2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6870
      RightToLeft     =   -1  'True
      TabIndex        =   443
      Top             =   2280
      Width           =   795
   End
   Begin VB.TextBox txtTradingContractID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3720
      TabIndex        =   442
      TabStop         =   0   'False
      Top             =   2280
      Width           =   825
   End
   Begin VB.Frame fra 
      Caption         =   "ÿ·» ’—› „ ⁄ÂœÌ‰"
      Height          =   975
      Index           =   4
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   425
      Top             =   900
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox TxtOrderSuppler 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3360
         TabIndex        =   428
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "⁄—÷"
         Height          =   315
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   427
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "«·€«¡ «·”œ«œ"
         Height          =   315
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   426
         Top             =   240
         Width           =   975
      End
      Begin MSDataListLib.DataCombo DcDur 
         Height          =   315
         Left            =   1800
         TabIndex        =   429
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcMontth 
         Height          =   315
         Left            =   120
         TabIndex        =   430
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbBrReq 
         Height          =   315
         Left            =   120
         TabIndex        =   431
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·”‰… "
         Height          =   225
         Index           =   72
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   435
         Top             =   600
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·‘Â—"
         Height          =   240
         Index           =   71
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   434
         Top             =   615
         Width           =   690
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·ÿ·»"
         Height          =   195
         Index           =   55
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   433
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·›—⁄"
         Height          =   195
         Index           =   73
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   432
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.TextBox TxtPrePayd 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   315
      Index           =   17
      Left            =   9000
      MaxLength       =   10
      TabIndex        =   424
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox TxtCurrencyRate 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10230
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   391
      Top             =   2250
      Width           =   615
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E2E9E9&
      Height          =   1005
      Index           =   2
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   323
      Top             =   2520
      Width           =   5835
      Begin VB.TextBox TxtAccount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   324
         Top             =   240
         Width           =   705
      End
      Begin MSDataListLib.DataCombo DcbAccount 
         Height          =   315
         Left            =   120
         TabIndex        =   325
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Õ”«»"
         Height          =   285
         Index           =   91
         Left            =   4800
         TabIndex        =   326
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.TextBox txtperson 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   254
      Top             =   4560
      Width           =   3285
   End
   Begin VB.TextBox TxtEndService 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   212
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "„’«—Ì› ÕÊ«·… »‰ﬂÌ…"
      Height          =   495
      Left            =   11460
      RightToLeft     =   -1  'True
      TabIndex        =   192
      Top             =   9300
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "—”Ê„ «·ÕÊ«·Â"
         Height          =   255
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   193
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox TxtReportName 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   181
      Top             =   2160
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CboPaymentType 
      Height          =   315
      ItemData        =   "FrmPayments.frx":0CE6
      Left            =   6360
      List            =   "FrmPayments.frx":0CE8
      RightToLeft     =   -1  'True
      TabIndex        =   179
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·„ÕÊ·"
      Height          =   1335
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   172
      Top             =   3360
      Width           =   6735
      Begin VB.TextBox TxtAdress2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox TxtCountry 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1290
         RightToLeft     =   -1  'True
         TabIndex        =   187
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox TxtGovernorate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   186
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox TxtCity 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1290
         RightToLeft     =   -1  'True
         TabIndex        =   185
         Top             =   600
         Width           =   1125
      End
      Begin VB.TextBox TxtStreet 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   183
         Top             =   600
         Width           =   1125
      End
      Begin VB.TextBox TxtRemitterName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3600
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtNumIqama 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3600
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TxtTelephone 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3600
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„œÌ‰…/«·‘«—⁄"
         Height          =   285
         Index           =   63
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   190
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄‰Ê«‰"
         Height          =   285
         Index           =   62
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   189
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "œÊ·…/«·„Õ«›Ÿ…"
         Height          =   285
         Index           =   61
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   184
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÕÊ·"
         Height          =   285
         Index           =   58
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   178
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÂÊÌ… «·„Êœ⁄"
         Height          =   285
         Index           =   56
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   176
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ·›Ê‰ «·„Êœ⁄"
         Height          =   285
         Index           =   57
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   175
         Top             =   960
         Width           =   1155
      End
   End
   Begin VB.TextBox TxtNoSupplerDes 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2400
      TabIndex        =   171
      TabStop         =   0   'False
      Top             =   1650
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtDue 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   168
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox PayDes 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5160
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   167
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox empDes1 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   162
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox empDes 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   158
      Top             =   4080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox TxtManulaNO 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9180
      RightToLeft     =   -1  'True
      TabIndex        =   152
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox XPTxtValView 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      Left            =   6540
      TabIndex        =   138
      Top             =   1515
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxtAdvance 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Txtorder 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   600
      Width           =   1335
   End
   Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
      Height          =   315
      Left            =   6840
      TabIndex        =   131
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      _extentx        =   2566
      _extenty        =   556
   End
   Begin VB.TextBox XPTxtID1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15480
      RightToLeft     =   -1  'True
      TabIndex        =   125
      Text            =   "Text4"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„«  «·ÕÊ«·Â"
      Height          =   1815
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   107
      Top             =   9840
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   108
         Top             =   240
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   120
         TabIndex        =   109
         Top             =   570
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Format          =   238223361
         CurrentDate     =   39614
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒÂ«"
         Height          =   285
         Index           =   39
         Left            =   2190
         RightToLeft     =   -1  'True
         TabIndex        =   111
         Top             =   570
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÕÊ«·Â"
         Height          =   285
         Index           =   38
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   110
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox TxtCustCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12420
      RightToLeft     =   -1  'True
      TabIndex        =   106
      Top             =   1920
      Width           =   1185
   End
   Begin VB.TextBox txt_ORDER_NO 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4410
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txt_general_des 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   10680
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   8490
      Width           =   3135
   End
   Begin VB.TextBox txtNoteSerial1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12435
      RightToLeft     =   -1  'True
      TabIndex        =   98
      Top             =   600
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   96
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAdv_payment_value 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   15060
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   94
      Top             =   2595
      Width           =   2685
   End
   Begin VB.Frame fra 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ŒÌ«—« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   88
      Top             =   480
      Width           =   3735
      Begin VB.Frame Frame18 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   437
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
         Begin VB.OptionButton Option8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "›Ê« Ì— „‘ —Ì« "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   0
            Left            =   120
            MaskColor       =   &H00000000&
            RightToLeft     =   -1  'True
            TabIndex        =   439
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton Option8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "›Ê« Ì— „«·Ì…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            MaskColor       =   &H00000000&
            RightToLeft     =   -1  'True
            TabIndex        =   438
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "⁄—÷ «·„” Œ·’« "
         Height          =   315
         Index           =   2
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   330
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "⁄—÷ «·›Ê« Ì— «·„«·Ì…"
         Height          =   315
         Index           =   0
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   290
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "⁄—÷ ›Ê« Ì— «·„‘ —Ì« "
         Height          =   315
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   289
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "œ›⁄Â „ﬁœ„Â"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "FIFO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   " ÕœÌœ ›Ê« Ì—"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin ALLButtonS.ALLButton ALLButton3 
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   " ÕœÌœ"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmPayments.frx":0CEA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame FraNote 
      BackColor       =   &H00E2E9E9&
      Height          =   1245
      Left            =   10530
      RightToLeft     =   -1  'True
      TabIndex        =   71
      Top             =   3360
      Width           =   4155
      Begin VB.TextBox TxtChequeNumber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1950
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   810
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker DtpChequeDueDate 
         Height          =   315
         Left            =   30
         TabIndex        =   6
         Top             =   840
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   242417665
         CurrentDate     =   39614
      End
      Begin MSDataListLib.DataCombo DcboBankName 
         Height          =   315
         Left            =   30
         TabIndex        =   4
         Top             =   480
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboBox 
         Height          =   315
         Left            =   30
         TabIndex        =   3
         Top             =   150
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ"
         Height          =   285
         Index           =   17
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   191
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·Œ“Ì‰Â"
         Height          =   285
         Index           =   9
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·»‰ﬂ"
         Height          =   285
         Index           =   15
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   510
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·‘Ìﬂ"
         Height          =   285
         Index           =   16
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   810
         Width           =   1215
      End
   End
   Begin VB.Frame FraInfo 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„«   Â„ﬂ"
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
      Height          =   2265
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   2400
      Width           =   3705
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   0
         Left            =   1830
         TabIndex        =   51
         Top             =   780
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments.frx":0D06
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   52
         Top             =   780
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments.frx":0E68
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   2
         Left            =   1830
         TabIndex        =   53
         Top             =   1350
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments.frx":0FCA
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   54
         Top             =   1350
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments.frx":112C
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   4
         Left            =   1830
         TabIndex        =   55
         Top             =   1920
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments.frx":128E
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   56
         Top             =   1920
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments.frx":13F0
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   57
         Top             =   540
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments.frx":1552
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   58
         Top             =   1110
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments.frx":16B4
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   59
         Top             =   1680
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments.frx":1816
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„œ›Ê⁄«  ›Ï «·≈”»Ê⁄ «·Õ«·Ï:"
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
         Height          =   225
         Index           =   19
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   1110
         Width           =   2235
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„œ›Ê⁄«  ›Ï «·‘Â— «·Õ«·Ï :"
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
         Height          =   225
         Index           =   20
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   1680
         Width           =   2235
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ﬁœÌ"
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
         Height          =   225
         Index           =   21
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·≈”»Ê⁄ «·Õ«·Ï"
         Height          =   255
         Index           =   22
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈Ã„«·Ï „œ›Ê⁄«  «·ÌÊ„:"
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
         Height          =   225
         Index           =   23
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   540
         Width           =   2235
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Ìﬂ« "
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
         Height          =   225
         Index           =   24
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   1350
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ﬁœÌ"
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
         Height          =   225
         Index           =   25
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Ìﬂ« "
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
         Height          =   225
         Index           =   26
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ﬁœÌ"
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
         Height          =   225
         Index           =   27
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   780
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Ìﬂ« "
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
         Height          =   225
         Index           =   28
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   780
         Width           =   675
      End
   End
   Begin VB.CheckBox ChkTrans 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰ Õ”«» ›« Ê—…"
      Height          =   225
      Left            =   14970
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   390
      Width           =   1575
   End
   Begin VB.Frame fra 
      BackColor       =   &H00E2E9E9&
      Height          =   975
      Index           =   0
      Left            =   15510
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   630
      Width           =   3675
      Begin VB.TextBox TxtTransID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   180
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TxtTransSerial 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   540
         Width           =   1005
      End
      Begin VB.ComboBox CboTrans 
         Height          =   315
         Left            =   600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   210
         Width           =   1995
      End
      Begin ImpulseButton.ISButton CmdSearchTrans 
         Height          =   345
         Left            =   600
         TabIndex        =   32
         Top             =   540
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
         ButtonImage     =   "FrmPayments.frx":1978
      End
      Begin ImpulseButton.ISButton CmdOpenTrans 
         Height          =   345
         Left            =   90
         TabIndex        =   34
         Top             =   540
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
         ButtonImage     =   "FrmPayments.frx":1D12
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
         TabIndex        =   36
         Top             =   600
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
         TabIndex        =   35
         Top             =   270
         Width           =   1305
      End
   End
   Begin VB.ComboBox DCboCashType 
      Height          =   315
      Left            =   11700
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   1905
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   465
      Left            =   10710
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   8010
      Width           =   3105
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   12060
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2250
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   6990
      TabIndex        =   22
      Top             =   8460
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   11070
      TabIndex        =   37
      Top             =   8850
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   10185
      TabIndex        =   38
      Top             =   8850
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   9300
      TabIndex        =   39
      Top             =   8850
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
      Left            =   8415
      TabIndex        =   40
      Top             =   8850
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   7530
      TabIndex        =   41
      Top             =   8850
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   2520
      TabIndex        =   42
      Top             =   8850
      Width           =   855
      _ExtentX        =   1508
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   14
      Left            =   3480
      TabIndex        =   43
      Top             =   8850
      Width           =   855
      _ExtentX        =   1508
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
      Left            =   6645
      TabIndex        =   44
      Top             =   8880
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
      Index           =   7
      Left            =   5670
      TabIndex        =   45
      Top             =   8880
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
   Begin ImpulseAniLabel.ISAniLabel LblLink 
      Height          =   315
      Left            =   60
      TabIndex        =   47
      Top             =   1635
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   556
      ActiveUnderline =   -1  'True
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ForeColor       =   4210688
      MousePointer    =   99
      MouseIcon       =   "FrmPayments.frx":20AC
      BackColor       =   14871017
      Alignment       =   1
      Caption         =   ""
      ColorHover      =   16711680
      RightToLeft     =   -1  'True
      ImageCount      =   0
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   48
      Top             =   9000
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
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
   Begin MSDataListLib.DataCombo DCPROJECT 
      Height          =   315
      Left            =   15240
      TabIndex        =   86
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcCostCenter 
      Bindings        =   "FrmPayments.frx":220E
      Height          =   315
      Left            =   11160
      TabIndex        =   10
      Top             =   2640
      Width           =   2445
      _ExtentX        =   4313
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   10
      Left            =   7200
      TabIndex        =   97
      Top             =   9240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   9
      Left            =   6000
      TabIndex        =   100
      Top             =   9240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
   Begin MSDataListLib.DataCombo DcBranch 
      Height          =   315
      Left            =   6960
      TabIndex        =   1
      Top             =   960
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   15
      Left            =   3840
      TabIndex        =   132
      Top             =   9240
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   12
      Left            =   4440
      TabIndex        =   133
      Top             =   8850
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "‰”Œ… „„«À·Â"
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
   Begin MSDataListLib.DataCombo DCPROJECT1 
      Height          =   315
      Left            =   11160
      TabIndex        =   139
      Top             =   3000
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   556
      _Version        =   393216
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
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   7050
      TabIndex        =   142
      Top             =   600
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      Format          =   242417665
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo Dcterm1 
      Height          =   315
      Left            =   3720
      TabIndex        =   143
      Top             =   2640
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo dcopr 
      Height          =   315
      Left            =   3720
      TabIndex        =   145
      Top             =   3000
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   11685
      TabIndex        =   154
      Top             =   600
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   8880
      TabIndex        =   180
      Top             =   1920
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   13
      Left            =   4800
      TabIndex        =   182
      Top             =   9240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… «·«Ìœ«⁄"
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
   Begin VB.TextBox txtTransferExpenses 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8880
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   291
      Top             =   2640
      Width           =   975
   End
   Begin MSDataListLib.DataCombo DcbEmpBranch 
      Height          =   315
      Left            =   4440
      TabIndex        =   327
      Top             =   1320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbContractor 
      Height          =   315
      Left            =   4440
      TabIndex        =   328
      Top             =   1560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbCurrency 
      Height          =   315
      Left            =   4440
      TabIndex        =   395
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox IncludVAT 
      Height          =   270
      Left            =   8760
      TabIndex        =   420
      Top             =   3000
      Width           =   2175
      _Version        =   786432
      _ExtentX        =   3836
      _ExtentY        =   476
      _StockProps     =   79
      Caption         =   "«·ÕÊ«·…  ‘„· «·ﬁÌ„… «·„÷«›…"
      ForeColor       =   8388608
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.TextBox XPTxtValE 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   8160
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   422
      Top             =   2280
      Width           =   1305
   End
   Begin VB.Frame Frame7 
      Caption         =   "„œ›Ê⁄«  „ﬁœ„…"
      Height          =   495
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   164
      Top             =   960
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton Command3 
         Caption         =   "«·€«¡ «·”œ«œ"
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   166
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "⁄—÷"
         Height          =   315
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   165
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "—Ê« » ⁄‰"
      Height          =   975
      Left            =   11160
      RightToLeft     =   -1  'True
      TabIndex        =   155
      Top             =   1740
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "«·€«¡ «·”œ«œ"
         Height          =   345
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   163
         Top             =   510
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "⁄—÷"
         Height          =   285
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   159
         Top             =   540
         Width           =   645
      End
      Begin VB.ComboBox CboYear1 
         Height          =   315
         Left            =   360
         TabIndex        =   157
         Text            =   "CboYear1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox CmbMonth1 
         Height          =   315
         Left            =   1920
         TabIndex        =   156
         Text            =   "CmbMonth1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "”‰…"
         Height          =   255
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   161
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "‘Â—"
         Height          =   255
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   160
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì Õ«·… «·„ÊŸ›"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10350
      RightToLeft     =   -1  'True
      TabIndex        =   101
      Top             =   1350
      Width           =   4455
      Begin VB.OptionButton Option7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»œ·«  „ﬁœ„Â"
         Height          =   195
         Left            =   -480
         RightToLeft     =   -1  'True
         TabIndex        =   129
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Œ’’« "
         Height          =   195
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   128
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«ÃÊ— „” Õﬁ…"
         Height          =   195
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”·›…"
         Height          =   195
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   102
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì Õ«·… „ﬁ«Ê·Ì «·»«ÿ‰"
      Height          =   495
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   147
      Top             =   1230
      Visible         =   0   'False
      Width           =   4215
      Begin VB.OptionButton subContOpt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "÷„«‰ «⁄„«·"
         Height          =   195
         Index           =   1
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   150
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton subContOpt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "œ›⁄«  „ﬁœ„…"
         Height          =   195
         Index           =   0
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   149
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton subContOpt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«⁄„«·"
         Height          =   195
         Index           =   2
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   148
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "»œ·«  „ﬁœ„…"
      Height          =   495
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   417
      Top             =   1320
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Command11 
         Caption         =   "«·€«¡ «·”œ«œ"
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   419
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command12 
         Caption         =   "⁄—÷"
         Height          =   315
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   418
         Top             =   120
         Width           =   975
      End
   End
   Begin MSDataListLib.DataCombo DcbDepartment 
      Height          =   315
      Left            =   6360
      TabIndex        =   436
      Top             =   1920
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   585
      Left            =   -30
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   14745
      _cx             =   26009
      _cy             =   1032
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
      BackColor       =   12648447
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "”‰œ ’—› - «·„œ›Ê⁄«   "
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
      FrameColor      =   8454143
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   440
         Top             =   0
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox oldtxtNoteSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   126
         Top             =   120
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   2940
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   120
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   3420
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1155
         TabIndex        =   12
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
         ButtonImage     =   "FrmPayments.frx":2223
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
         Left            =   90
         TabIndex        =   13
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
         ButtonImage     =   "FrmPayments.frx":25BD
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
         Left            =   1680
         TabIndex        =   14
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
         ButtonImage     =   "FrmPayments.frx":2957
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
         Left            =   615
         TabIndex        =   15
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
         ButtonImage     =   "FrmPayments.frx":2CF1
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin MSAdodcLib.Adodc detect_no 
         Height          =   585
         Left            =   120
         Top             =   120
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
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   6360
         Picture         =   "FrmPayments.frx":308B
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "»Ì«‰«  ‰Â«Ì… «·Œœ„…"
      Height          =   3855
      Index           =   3
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   194
      Top             =   4530
      Visible         =   0   'False
      Width           =   13215
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   16
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   387
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   15
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   386
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   14
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   385
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   13
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   384
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   12
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   383
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   11
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   321
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   10
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   319
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   9
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   317
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   315
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   313
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   312
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   8400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   308
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   8400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   307
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   8400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   306
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   8400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   305
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   8400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   303
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox TxtPrePayd 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   8400
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   301
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   299
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtTotlPaidEndSer 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   298
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtCusTiket 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   297
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtTicktConract 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   295
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox TxtValEndService2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   287
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtTotalDis 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   284
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox TxtTotalDis2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   283
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox TxtAddOther 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   280
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox TxtAddOther2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   279
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox TxtCusTiket2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   277
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtValEndService 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   275
         Text            =   "0"
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TXTLastTotal2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   271
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtTicketValue2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   266
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtCustom2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   265
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox TxtCash2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   264
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TXTAdvanceTotal2 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   263
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtSal2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   259
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtnet2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   10440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   258
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txttotal2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   10440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   257
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtVlueVaction2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   4080
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   256
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10920
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   214
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TXTLastTotal 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   210
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox TXTAdvanceTotal 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   206
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TxtCash 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   205
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TxtVlueVaction 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   204
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtnet 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   198
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtCustom 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   197
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtTicketValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   196
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtSal 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   195
         Top             =   1320
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DcbEmpEndService 
         Height          =   315
         Left            =   6480
         TabIndex        =   215
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         ListField       =   ""
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbBranchEndServ 
         Height          =   315
         Left            =   120
         TabIndex        =   217
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   40
         Left            =   1320
         TabIndex        =   390
         Top             =   2520
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   285
         Index           =   34
         Left            =   3000
         TabIndex        =   389
         Top             =   2520
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’Ê„«  —« »"
         Height          =   285
         Index           =   15
         Left            =   4920
         TabIndex        =   388
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   285
         Index           =   39
         Left            =   3000
         TabIndex        =   322
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   285
         Index           =   38
         Left            =   3000
         TabIndex        =   320
         Top             =   2880
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   285
         Index           =   37
         Left            =   3000
         TabIndex        =   318
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   285
         Index           =   36
         Left            =   3000
         TabIndex        =   316
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   285
         Index           =   35
         Left            =   3000
         TabIndex        =   314
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   285
         Index           =   33
         Left            =   9360
         TabIndex        =   311
         Top             =   2880
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   285
         Index           =   32
         Left            =   9360
         TabIndex        =   310
         Top             =   2520
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   285
         Index           =   31
         Left            =   9360
         TabIndex        =   309
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   285
         Index           =   30
         Left            =   9360
         TabIndex        =   304
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   285
         Index           =   28
         Left            =   9360
         TabIndex        =   302
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   0
         Left            =   7800
         TabIndex        =   300
         Top             =   840
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «· –«ﬂ— ÿ»ﬁ« ··⁄ﬁœ"
         Height          =   285
         Index           =   29
         Left            =   11520
         TabIndex        =   296
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   27
         Left            =   1320
         TabIndex        =   294
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’Ê„«  «Œ—Ï"
         Height          =   285
         Index           =   25
         Left            =   4920
         TabIndex        =   293
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   24
         Left            =   1320
         TabIndex        =   288
         Top             =   840
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì «·Œ’Ê„« "
         Height          =   285
         Index           =   23
         Left            =   4920
         TabIndex        =   286
         Top             =   2880
         Width           =   1515
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   22
         Left            =   1320
         TabIndex        =   285
         Top             =   2880
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "≈÷«›«  «Œ—Ï"
         Height          =   285
         Index           =   21
         Left            =   11760
         TabIndex        =   282
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   20
         Left            =   7800
         TabIndex        =   281
         Top             =   2880
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… „Œ’’ «· –«ﬂ—"
         Height          =   285
         Index           =   16
         Left            =   11760
         TabIndex        =   278
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Œ’Ê„« "
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   26
         Left            =   4920
         TabIndex        =   276
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   12360
         RightToLeft     =   -1  'True
         TabIndex        =   274
         Top             =   240
         Width           =   495
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "’«›Ì «·„œ›Ê⁄"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   14
         Left            =   600
         TabIndex        =   272
         Top             =   3240
         Width           =   1755
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   9
         Left            =   1320
         TabIndex        =   270
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   8
         Left            =   1320
         TabIndex        =   269
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   6
         Left            =   7800
         TabIndex        =   268
         Top             =   1680
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   4
         Left            =   7800
         TabIndex        =   267
         Top             =   2520
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   262
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   2
         Left            =   7800
         TabIndex        =   261
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄"
         Height          =   285
         Index           =   1
         Left            =   7800
         TabIndex        =   260
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "«·›—⁄"
         Height          =   285
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   218
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„ÊŸ›"
         Height          =   285
         Left            =   10920
         RightToLeft     =   -1  'True
         TabIndex        =   216
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "’«›Ì «·„ﬂ«›√…"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   12
         Left            =   1200
         TabIndex        =   211
         Top             =   4080
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì «·”·›"
         Height          =   285
         Index           =   10
         Left            =   4920
         TabIndex        =   209
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì «·⁄Âœ"
         Height          =   285
         Index           =   11
         Left            =   4920
         TabIndex        =   208
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«Ã«“«  »œÊ‰ —« »"
         Height          =   285
         Index           =   13
         Left            =   4920
         TabIndex        =   207
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Â«Ì… «·Œœ„…"
         Height          =   285
         Index           =   5
         Left            =   11760
         TabIndex        =   203
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì «·«÷«›« "
         Height          =   285
         Index           =   7
         Left            =   11760
         TabIndex        =   202
         Top             =   3240
         Width           =   1395
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «· –ﬂ—…"
         Height          =   285
         Index           =   17
         Left            =   11760
         TabIndex        =   201
         Top             =   3840
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„Œ’’ «·«Ã«“…"
         Height          =   285
         Index           =   18
         Left            =   11760
         TabIndex        =   200
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—« » «·‘Â— «·Õ«·Ï"
         Height          =   285
         Index           =   19
         Left            =   11760
         TabIndex        =   199
         Top             =   1320
         Width           =   1395
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "»Ì«‰«  «” Õﬁ«ﬁ «·«Ã«“…"
      Height          =   3735
      Left            =   210
      RightToLeft     =   -1  'True
      TabIndex        =   331
      Top             =   4080
      Visible         =   0   'False
      Width           =   9735
      Begin VB.TextBox txtPaymentRecommended 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   449
         TabStop         =   0   'False
         Top             =   3330
         Width           =   1215
      End
      Begin VB.TextBox txtSalary 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   355
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtSalEntitOther 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   354
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtOther 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   353
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtSalaryVocation 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   352
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtValueTickt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   351
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtAdvance1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   350
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox TxtTotalsalary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   349
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7320
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   348
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TxtInsuranceValue 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   347
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox TxtInsuranceValue2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   346
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtAdvance12 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   345
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox TxtOther2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   344
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox TxtSalEntitOther2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   343
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtSalary2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   342
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtValueTickt2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   341
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtSalaryVocation2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   340
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox TxtTotalsalary2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   339
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtSalaryVocation3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   338
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtValueTickt3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   337
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtSalary3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   336
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtSalEntitOther3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   335
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtOther3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   334
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtAdvance13 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   333
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox TxtInsuranceValue3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   332
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DcBranch1 
         Height          =   315
         Left            =   120
         TabIndex        =   356
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   3960
         TabIndex        =   357
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         ListField       =   ""
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ê’Ï »”œ«œ"
         Height          =   255
         Index           =   100
         Left            =   8400
         TabIndex        =   448
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "—« » «·‘Â— «·Õ«·Ì"
         Height          =   255
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   382
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "«Œ—Ì «÷«›« "
         Height          =   255
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   381
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "«Œ—Ì Œ’Ê„« "
         Height          =   255
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   380
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "”·› ”«»ﬁ…"
         Height          =   255
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   379
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   378
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   " –«ﬂ— «·”›—"
         Height          =   255
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   377
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "„” Õﬁ«  «·«Ã«“…"
         Height          =   255
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   376
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„ÊŸ›"
         Height          =   255
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   375
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   374
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÌ„… «· √„Ì‰"
         Height          =   255
         Index           =   74
         Left            =   8400
         TabIndex        =   373
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„œ›Ê⁄"
         Height          =   255
         Index           =   75
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   372
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„œ›Ê⁄"
         Height          =   255
         Index           =   76
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   371
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„œ›Ê⁄"
         Height          =   255
         Index           =   77
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   370
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„œ›Ê⁄"
         Height          =   255
         Index           =   78
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   369
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„œ›Ê⁄"
         Height          =   255
         Index           =   79
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   368
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„œ›Ê⁄"
         Height          =   255
         Index           =   80
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   367
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„œ›Ê⁄"
         Height          =   255
         Index           =   81
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   366
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label40 
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
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   365
         Top             =   120
         Width           =   135
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   255
         Index           =   82
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   364
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   255
         Index           =   83
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   363
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   255
         Index           =   84
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   362
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   255
         Index           =   85
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   361
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   255
         Index           =   86
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   360
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   255
         Index           =   87
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   359
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "„œ›Ê⁄ „”»ﬁ«"
         Height          =   255
         Index           =   88
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   358
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
      Height          =   2595
      Index           =   2
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   113
      Top             =   4800
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox TxtPaymentCounts 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   3750
         MaxLength       =   2
         TabIndex        =   118
         Top             =   240
         Width           =   825
      End
      Begin VB.ComboBox CmbMonth 
         Height          =   315
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   117
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox ChkSaleryDis 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈Œ’„ „‰ «·„— »  ·ﬁ«∆Ì«"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3720
         TabIndex        =   116
         Top             =   2160
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.ComboBox CboYear 
         Height          =   315
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   1320
         Width           =   1095
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   435
         Index           =   11
         Left            =   3750
         TabIndex        =   114
         Top             =   1680
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   767
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "≈Õ”»  Ê«—ÌŒ «·”œ«œ"
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
         ButtonImage     =   "FrmPayments.frx":6CF3
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   1845
         Left            =   90
         TabIndex        =   119
         Top             =   210
         Width           =   3495
         _cx             =   6165
         _cy             =   3254
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPayments.frx":708D
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
      Begin VB.Label LblTotalV 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   130
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·œ›⁄« "
         Height          =   285
         Index           =   44
         Left            =   4470
         TabIndex        =   124
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «Ê· œ›⁄…"
         Height          =   285
         Index           =   43
         Left            =   3660
         TabIndex        =   123
         Top             =   690
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Ì„ﬂ‰ﬂ «· ⁄œÌ· ›Ï ﬁÌ„… «·œ›⁄«  ÌœÊÌ«ı"
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
         Height          =   255
         Left            =   60
         TabIndex        =   122
         Top             =   2280
         Width           =   2595
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Â—"
         Height          =   315
         Index           =   42
         Left            =   4890
         TabIndex        =   121
         Top             =   990
         Width           =   405
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”‰…"
         Height          =   315
         Index           =   41
         Left            =   4890
         TabIndex        =   120
         Top             =   1320
         Width           =   405
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E2E9E9&
      Height          =   2805
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   236
      Top             =   4800
      Width           =   4395
      Begin VB.TextBox TxtBeneficiaryACNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         TabIndex        =   246
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox TxtBenefiBanckCode 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1920
         TabIndex        =   245
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TxtBenefiBanckAddress 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   244
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox TxtBenefStreet 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   243
         Top             =   2040
         Width           =   1485
      End
      Begin VB.TextBox TxtBenefCity 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1650
         RightToLeft     =   -1  'True
         TabIndex        =   242
         Top             =   2040
         Width           =   1365
      End
      Begin VB.TextBox TxtBenefTelephone 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         TabIndex        =   241
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TxtBenefIBAN 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         TabIndex        =   240
         TabStop         =   0   'False
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox TxtKafelAddress 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   239
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox TxtBeneficiaryBanck 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   237
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label38 
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
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   273
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄‰Ê«‰ »‰ﬂ «·„” ›Ìœ"
         Height          =   255
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   253
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—„“ »‰ﬂ «·„” ›Ìœ"
         Height          =   255
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   252
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ Õ”«» «·„” ›Ìœ"
         Height          =   255
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   251
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ·›Ê‰"
         Height          =   255
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   250
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «Ì»«‰ «·„” ›Ìœ"
         Height          =   255
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   249
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„œÌ‰…/«·‘«—⁄"
         Height          =   255
         Index           =   69
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   248
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄‰Ê«‰ «·ﬂ›Ì·"
         Height          =   255
         Index           =   66
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   247
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»‰ﬂ «·„” ›Ìœ"
         Height          =   255
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   238
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00E2E9E9&
      Height          =   2445
      Left            =   9720
      RightToLeft     =   -1  'True
      TabIndex        =   219
      Top             =   4920
      Width           =   4995
      Begin VB.TextBox TxtBenefGovernorate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   227
         Top             =   1680
         Width           =   1485
      End
      Begin VB.TextBox TXtBenefCountry 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   226
         Top             =   1680
         Width           =   1845
      End
      Begin VB.TextBox TxtBeneficiaryAddress 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   240
         Width           =   3405
      End
      Begin VB.TextBox TxtBenefNumIqama 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   224
         Top             =   600
         Width           =   3405
      End
      Begin VB.TextBox TxtBenefPlaceBrith 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   223
         Top             =   1320
         Width           =   1845
      End
      Begin VB.TextBox TxtBenefPlaceIqama 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   222
         Top             =   960
         Width           =   1845
      End
      Begin VB.TextBox TxtKafeltEL 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   221
         Top             =   2040
         Width           =   1485
      End
      Begin VB.TextBox TxtKafelName 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   220
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker BenefBrithDate 
         Height          =   315
         Left            =   120
         TabIndex        =   228
         Top             =   1320
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         Format          =   245825537
         CurrentDate     =   39614
      End
      Begin Dynamic_Byte.NourHijriCal BenefDateExpEqama 
         Height          =   315
         Left            =   120
         TabIndex        =   229
         Top             =   960
         Width           =   1485
         _extentx        =   2619
         _extenty        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄‰Ê«‰ «·„” ›Ìœ"
         Height          =   285
         Index           =   97
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   235
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«ﬁ«„…/«·ÂÊÌ…"
         Height          =   285
         Index           =   59
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   234
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "œÊ·…/«·„Õ«›Ÿ…"
         Height          =   285
         Index           =   60
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   233
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ﬂ«‰/ «—ÌŒ «·«’œ«—"
         Height          =   285
         Index           =   67
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   232
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„ﬂ«‰/ «—ÌŒ «·„Ì·«œ"
         Height          =   285
         Index           =   68
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   231
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬂ›Ì·/ ·›Ê‰Â"
         Height          =   285
         Index           =   64
         Left            =   3570
         RightToLeft     =   -1  'True
         TabIndex        =   230
         Top             =   2040
         Width           =   1335
      End
   End
   Begin VB.Frame fra 
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
      Left            =   -90
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   7560
      Width           =   10095
      Begin VB.TextBox TxtValueTemp 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5880
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   170
         Top             =   240
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   240
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   76
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
         TabIndex        =   77
         Top             =   510
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› „œÌ‰"
         Height          =   285
         Index           =   32
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   180
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› œ«∆‰"
         Height          =   285
         Index           =   31
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   510
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   315
         Index           =   30
         Left            =   8970
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   210
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·› —… :"
         Height          =   315
         Index           =   29
         Left            =   8970
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   540
         Width           =   975
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   11
         Left            =   8670
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   510
         Width           =   1485
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„œ…"
      Height          =   285
      Index           =   101
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   450
      Top             =   1950
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«Ã„«·Ì"
      Height          =   285
      Index           =   99
      Left            =   6060
      RightToLeft     =   -1  'True
      TabIndex        =   446
      Top             =   2310
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Vat"
      Height          =   285
      Index           =   98
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   444
      Top             =   2310
      Width           =   405
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·« ›«ﬁÌ…"
      Height          =   285
      Index           =   95
      Left            =   4530
      TabIndex        =   441
      Top             =   2310
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌ„… «·„÷«›…"
      Height          =   345
      Index           =   65
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   423
      Top             =   1320
      Width           =   780
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„—«ﬂ“ «· ﬂ·›…"
      Height          =   285
      Index           =   94
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   421
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·⁄„·…"
      Height          =   285
      Index           =   93
      Left            =   5520
      TabIndex        =   394
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ﬁ«»·"
      Height          =   255
      Index           =   96
      Left            =   9450
      RightToLeft     =   -1  'True
      TabIndex        =   393
      ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„⁄œ·"
      Height          =   285
      Index           =   92
      Left            =   11640
      RightToLeft     =   -1  'True
      TabIndex        =   392
      Top             =   2280
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ﬁ«Ê·"
      Height          =   285
      Index           =   90
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   329
      Top             =   1560
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "„’«—Ì› «·ÕÊ«·Â"
      Height          =   195
      Index           =   89
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   292
      Top             =   2640
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„” ›Ìœ"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   34
      Left            =   13860
      RightToLeft     =   -1  'True
      TabIndex        =   255
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "”‰œ ‰Â«Ì… «·Œœ„…"
      Height          =   285
      Index           =   70
      Left            =   5790
      TabIndex        =   213
      Top             =   1350
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "”‰œ «·„” Õﬁ« "
      Height          =   285
      Index           =   54
      Left            =   5670
      TabIndex        =   169
      Top             =   1350
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
      Height          =   285
      Index           =   53
      Left            =   10560
      RightToLeft     =   -1  'True
      TabIndex        =   153
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈œ«—…"
      Height          =   285
      Index           =   52
      Left            =   8250
      RightToLeft     =   -1  'True
      TabIndex        =   151
      Top             =   1950
      Width           =   555
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·⁄„·Ì…"
      Height          =   315
      Index           =   51
      Left            =   7935
      RightToLeft     =   -1  'True
      TabIndex        =   146
      Top             =   3000
      Width           =   750
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·»‰œ"
      Height          =   315
      Index           =   50
      Left            =   5895
      RightToLeft     =   -1  'True
      TabIndex        =   144
      Top             =   2640
      Width           =   390
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "›—⁄ «·„ÊŸ›"
      Height          =   315
      Index           =   49
      Left            =   8085
      RightToLeft     =   -1  'True
      TabIndex        =   141
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·„‘—Ê⁄"
      Height          =   285
      Index           =   48
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   140
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ÿ·» «·”·›Â"
      Height          =   285
      Index           =   47
      Left            =   5670
      TabIndex        =   137
      Top             =   1320
      Width           =   1365
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ ÿ·» «·’—›"
      Height          =   285
      Index           =   46
      Left            =   5670
      TabIndex        =   135
      Top             =   630
      Width           =   1365
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
      Height          =   195
      Index           =   45
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   127
      Top             =   9000
      Width           =   3675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   195
      Index           =   40
      Left            =   10860
      RightToLeft     =   -1  'True
      TabIndex        =   112
      Top             =   960
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»Ì…"
      Height          =   285
      Index           =   37
      Left            =   5310
      RightToLeft     =   -1  'True
      TabIndex        =   105
      Top             =   1950
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·‘—Õ «·⁄«„"
      Height          =   285
      Index           =   36
      Left            =   13560
      RightToLeft     =   -1  'True
      TabIndex        =   104
      Top             =   8490
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "œ›⁄Â „ﬁœ„Â"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   35
      Left            =   15720
      RightToLeft     =   -1  'True
      TabIndex        =   95
      Top             =   2655
      Width           =   1245
   End
   Begin VB.Label lblsqlstring 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   135
      Left            =   14400
      RightToLeft     =   -1  'True
      TabIndex        =   93
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„‘—Ê⁄"
      Height          =   285
      Index           =   33
      Left            =   14760
      RightToLeft     =   -1  'True
      TabIndex        =   87
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
      Height          =   315
      Index           =   14
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   2640
      Width           =   1245
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
      ForeColor       =   &H00000080&
      Height          =   405
      Index           =   18
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   2010
      Width           =   2715
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—’Ìœ «·Õ«·Ï:"
      Height          =   285
      Index           =   13
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   1650
      Width           =   1185
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Left            =   1830
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   8490
      Width           =   825
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   8490
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   255
      Index           =   6
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   8490
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   255
      Index           =   7
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   8490
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   300
      Index           =   8
      Left            =   9060
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   8490
      Width           =   1140
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„œ›Ê⁄« "
      Height          =   285
      Index           =   0
      Left            =   13680
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   615
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… «·„œ›Ê⁄« "
      Height          =   285
      Index           =   2
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2250
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„"
      Height          =   285
      Index           =   3
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1905
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”‰œ"
      Height          =   285
      Index           =   4
      Left            =   13680
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»‰«¡ ⁄·Ï"
      Height          =   285
      Index           =   5
      Left            =   13230
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   7980
      Width           =   1455
   End
End
Attribute VB_Name = "FrmPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim Line1 As Double
Dim Price As Double
Dim Line2 As Double
Dim Line3 As Double
Dim departement_name As Integer
Dim numbering_type As Integer
Dim Balance As String
Dim balanceString As String
Dim Account_Code_dynamic As String
Public called As Boolean
Dim FlgBill As Boolean
Dim FlgBillBuy As Boolean
Dim FlgBillProject As Boolean
Dim Ettfa As Boolean
Public EmpIDD As Double
Public ProjectIDD As Double
Dim Aut_manual As Boolean
 Dim currentname As String
 
Dim OtherInformation As New ClsGLOther
Function GetEmpID(Optional Fileds As String = "", Optional FiledWhere) As Double
GetEmpID = 0
If Fileds <> "" Then
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
sql = "Select Emp_ID from TblEmployee where " & Fileds & " ='" & FiledWhere & "'"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetEmpID = IIf(IsNull(Rs4("Emp_ID").value), 0, Rs4("Emp_ID").value)
Else
GetEmpID = 0
End If

End If
End Function







Private Sub Check1_Click()
    Dim i As Integer

    If Check1.value = vbChecked Then

        With Me.VSFlexGrid1
 
            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.VSFlexGrid1

            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If
    RelineBuy
End Sub
Private Sub Check2_Click()
    Dim i As Integer

    If Check2.value = vbChecked Then

        With Me.VSFlexGrid2
 
            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.VSFlexGrid2

            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If
    RelineProject
End Sub





Private Sub Command10_Click(Index As Integer)
Dim i As Integer
Dim StrSQL As String
If Index = 0 Then
If Me.TxtModFlg.text = "E" Then
DeleteBillBuy
VSFlexGrid1.Enabled = True
        Check1.Enabled = True
      StrSQL = "Delete From TblNotesBillBuyPayment Where NoteID1=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
XPTxtVal.text = 0
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
End If
If Index = 1 Then
If Me.TxtModFlg.text = "E" Then
DeleteBillProject
VSFlexGrid2.Enabled = True
        Check2.Enabled = True
      StrSQL = "Delete From TblNotesBillProjectPayment Where NoteID1=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillProjectPayment  Where NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
XPTxtVal.text = 0
            VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
VSFlexGrid2.rows = 1

FlgBillProject = True
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·€«¡ «·”œ«œ"
Else
MsgBox "Done"
End If
    With Me.VSFlexGrid2

            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i


        End With
End If
End If

End Sub

Private Sub Command11_Click()
'If ChekExpens(PayDes.text) = True Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'                MsgBox "·«Ì„ﬂ‰ ≈·€«¡ «·”œ«œ"
'        Else
'                 MsgBox "Can not Cancel "
'        End If
'        Exit Sub
'End If

XPTxtVal.text = 0
FrmEmpSalary6.ClearPayment = True
DeleteMofrdPayment PayDes.text, empDes1.text
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·€«¡ «·”œ«œ"
Else
MsgBox "Done"
End If
End Sub

Private Sub Command12_Click()
Unload FrmEmpSalary6
Load FrmEmpSalary6
FrmEmpSalary6.show
FrmEmpSalary6.Check22.Visible = True
FrmEmpSalary6.Frame3.Visible = False
FrmEmpSalary6.Command3.Visible = False
FrmEmpSalary6.Command4.Visible = False
FrmEmpSalary6.ALLButton8.Visible = False
FrmEmpSalary6.VSFlexGrid3.Visible = False
FrmEmpSalary6.Check21.Visible = False
FrmEmpSalary6.PayDes = PayDes.text
FrmEmpSalary6.PayDes1 = empDes1.text
FrmEmpSalary6.ALLButton3.Visible = False
FrmEmpSalary6.Grid2.Visible = False
FrmEmpSalary6.GRID1.Visible = False
FrmEmpSalary6.Check17.Visible = False
FrmEmpSalary6.lbl(12).Visible = False
FrmEmpSalary6.DTPicker1.Visible = False
FrmEmpSalary6.VSFlexGrid1.Visible = False
FrmEmpSalary6.Check19.Visible = False
FrmEmpSalary6.ALLButton7.Visible = False
FrmEmpSalary6.Check20.Visible = False
FrmEmpSalary6.VSFlexGrid2.Visible = False
FrmEmpSalary6.Check18.Visible = False
FrmEmpSalary6.ALLButton6.Visible = True
FrmEmpSalary6.VSFlexGrid4.Visible = True
FrmEmpSalary6.Check22.Visible = True
FrmEmpSalary6.FillGrid16
End Sub

Private Sub Command5_Click(Index As Integer)
If Index = 0 Then
Dim Msg As String
If val(DBCboClientName.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— «·„Ê—œ «Ê·«"
Else
MsgBox "Please Select Vendor"
End If
DBCboClientName.SetFocus
Exit Sub
Else
Aut_manual = False
Frame9.Visible = True
If Me.TxtModFlg.text <> "R" Then
XPTxtVal.text = 0
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "Â·  —Ìœ «· Ê“Ì⁄ ÌœÊÌ"
Else
Msg = "Do you want a manual distribution"
End If
 If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
Aut_manual = True
Else
Aut_manual = False
End If

If Me.TxtModFlg.text = "N" Then
RetriveBillVendor val(DBCboClientName.BoundText)
End If

If Me.TxtModFlg.text = "E" And (FlgBill = True Or GRID1.rows = 1) Then
RetriveBillVendor val(DBCboClientName.BoundText)
End If
End If
End If
End If
If Index = 1 Then

If val(DBCboClientName.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— «·„Ê—œ «Ê·«"
Else
MsgBox "Please Select Vendor"
End If
DBCboClientName.SetFocus
Exit Sub
Else
Frame12(0).Visible = True
If Me.TxtModFlg.text <> "R" Then
XPTxtVal.text = 0

If Me.TxtModFlg.text = "N" Then
    If DCboCashType.ListIndex = 13 Or DCboCashType.ListIndex = 14 Then
        RetriveBillBuy val(DBCboClientName.BoundText), val(txtTradingContractID)
    Else
        RetriveBillBuy val(DBCboClientName.BoundText)
    End If
End If

If Me.TxtModFlg.text = "E" And (FlgBillBuy = True Or VSFlexGrid1.rows = 1) Then
    If DCboCashType.ListIndex = 13 Or DCboCashType.ListIndex = 14 Then
        RetriveBillBuy val(DBCboClientName.BoundText), val(txtTradingContractID)
    Else
        RetriveBillBuy val(DBCboClientName.BoundText)
    End If
End If
End If
End If
End If

If Index = 2 Then

If val(DcbContractor.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«—  «·„ﬁ«Ê·"
Else
MsgBox "Please Select Contractor"
End If
DcbContractor.SetFocus
Exit Sub
Else
Frame12(1).Visible = True
If Me.TxtModFlg.text <> "R" Then
XPTxtVal.text = 0

If Me.TxtModFlg.text = "N" Then
RetriveBillProject val(DcbContractor.BoundText)
End If

If Me.TxtModFlg.text = "E" And (FlgBillProject = True Or VSFlexGrid2.rows = 1) Then
RetriveBillProject val(DcbContractor.BoundText)
End If
End If
End If
End If
End Sub




Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub

Private Sub DcbAccount_Click(Area As Integer)
TxtAccount.text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
'If Me.TxtModFlg.Text <> "R" Then
        If CboPayMentType.ListIndex = 4 Or CboPayMentType.ListIndex = 5 Then
            Me.DcboCreditSide.BoundText = DcbAccount.BoundText
        End If
' End If
End Sub

Private Sub DcbAccount_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 29125

    End If
End Sub

Private Sub DcbContractor_Change()
DcbContractor_Click (0)
End Sub

Private Sub DcbContractor_Click(Area As Integer)
If Me.TxtModFlg.text <> "R" Then
If val(DCboCashType.ListIndex) = 3 Then
If DBCboClientName.text = "" Or DBCboClientName.BoundText = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄ «Ê·«"
Else
MsgBox "Please Select Project"
End If
Exit Sub
End If
If val(DcbContractor.BoundText) <> 0 Then
Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbContractor.BoundText))
XPTxtVal.Enabled = False
Else
Me.DcboDebitSide.BoundText = GetProjectCoount(val((DBCboClientName.BoundText)))
XPTxtVal.Enabled = True
End If
End If
End If
End Sub

Private Sub DcbCurrency_Change()
DcbCurrency_Click (0)
End Sub

Private Sub DcbCurrency_Click(Area As Integer)
    If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    If Me.DcbCurrency.BoundText <> "" Then
        TxtCurrencyRate.text = get_currency_rate(Me.DcbCurrency.BoundText)
    Else
        TxtCurrencyRate.text = 1
    End If
CalCuteCurrency
End Sub

Private Sub DcboBankName_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
                    FrmExpensesSearch.Indx = 1
                    FrmExpensesSearch.show
                    FrmExpensesSearch.Indx = 1
                    FrmExpensesSearch.RetrunType = 2512
    End If
End Sub

Private Sub DcboBox_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
                    FrmExpensesSearch.Indx = 2
                    FrmExpensesSearch.show
                    FrmExpensesSearch.Indx = 2
                    FrmExpensesSearch.RetrunType = 2511
    End If
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcDur_Change()
Dim i As Integer, j As Integer, str As String
    i = val(dcDur.BoundText)
    
    If i > 0 Then
        str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo dcMontth, str
    Else
        str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo dcMontth, str
    End If

End Sub
'Private Sub ALLButton1_Click()

'    If IsNumeric(Me.DBCboClientName.BoundText) Then
'     '   INSTALLMENT_DATA2.show
'    ' '   INSTALLMENT_DATA2.Adodc1.CommandType = adCmdText
'     '   INSTALLMENT_DATA2.Adodc1.RecordSource = "select *  FROM INSTALLMENT_DETAILS where payed=1 and  cust_id =" & Me.DBCboClientName.BoundText
'     '   INSTALLMENT_DATA2.Adodc1.Refresh
' '
' '       INSTALLMENT_DATA2.id.text = Me.DBCboClientName.BoundText
' '       INSTALLMENT_DATA2.lblcustid = Me.DBCboClientName.BoundText
' '       INSTALLMENT_DATA2.TxtName.text = Me.DBCboClientName.text
'    End If

'End Sub

'Private Sub ALLButton2_Click()
'
'    If IsNumeric(Me.DBCboClientName.BoundText) Then
'        'sanad_dean.show
'        'sanad_dean.LblID = DBCboClientName.BoundText
'        'sanad_dean.LblName = DBCboClientName.text
'        ''sanad_dean.lblaccountcode.Caption = txtaccount.text
'        'sanad_dean.Adodc1.CommandType = adCmdText
''        sanad_dean.Adodc1.RecordSource = "select*  FROM sanad_dean where cust_id=" & DBCboClientName.BoundText
''        sanad_dean.Adodc1.Refresh
''        sanad_dean.ALLButton1.Visible = False
''        sanad_dean.ALLButton1.Visible = False
'
''        sanad_dean.Adodc2.CommandType = adCmdText
''        sanad_dean.Adodc2.RecordSource = "select *  FROM member_child where cust_id=" & DBCboClientName.BoundText
''        sanad_dean.Adodc2.Refresh
'    End If
'
'End Sub

Private Sub ALLButton3_Click()
    lblsqlstring.Caption = ""
    FrmPaymentTime2.show
    FrmPaymentTime2.lblcusid = val(DBCboClientName.BoundText)
    FrmPaymentTime2.LblValue = val(XPTxtVal.text)
    FrmPaymentTime2.LoadData val(DBCboClientName.BoundText)
End Sub

Private Sub CboPayMentType_Change()
Cmd(13).Enabled = False
Frame12(2).Visible = False
FraNote.Visible = True
txtTransferExpenses.Enabled = False
    If Me.TxtModFlg.text = "E" Then
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        Me.DcboBox.text = ""
        
    End If
    If Me.TxtModFlg.text <> "R" Then
    If val(Me.CboPayMentType.ListIndex) <> -1 Or val(Me.CboPayMentType.ListIndex) <> 0 Then
    If val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 1 Or val(DCboCashType.ListIndex) = 2 Then
    GetCustomer val(DBCboClientName.BoundText)
    End If
    If val(DCboCashType.ListIndex) = 4 Then
      GetEmployee EmpIDD
    End If
    End If
    End If
Frame11.Visible = False
Frame14.Visible = False
Frame15.Visible = False

    Frame4.Visible = False

    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(16).Caption = "—ﬁ„ «·‘Ìﬂ "
        lbl(17).Caption = " «—ÌŒ «·«” Õﬁ«ﬁ"
    
    Else
        lbl(16).Caption = "Cheque No"
        lbl(17).Caption = "Due Date"
    End If

    If Me.CboPayMentType.ListIndex = 0 Then
    txtTransferExpenses.text = 0
    Cmd(13).Enabled = False
        Me.lbl(9).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        DcbAccount.BoundText = ""
    ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 3 Then
    txtTransferExpenses.text = 0
    Cmd(13).Enabled = True
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Frame11.Visible = True
        Frame14.Visible = True
        Frame15.Visible = True
        DcbAccount.BoundText = ""
     ElseIf Me.CboPayMentType.ListIndex = 4 Then
     txtTransferExpenses.Enabled = True
     DcbAccount_Change
        Frame12(2).Visible = True
        FraNote.Visible = False
            Cmd(13).Enabled = False
        Frame11.Visible = False
        Frame14.Visible = False
        Frame15.Visible = False
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
          DcboBankName.BoundText = ""
        'TxtChequeNumber.text = ""

        Frame3.Enabled = True
        Frame4.Visible = True
        
 ElseIf Me.CboPayMentType.ListIndex = 5 Then
    Option3.value = True
     txtTransferExpenses.Enabled = True
     DcbAccount_Change
        Frame12(2).Visible = True
        FraNote.Visible = False
            Cmd(13).Enabled = False
        Frame11.Visible = False
        Frame14.Visible = False
        Frame15.Visible = False
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
          DcboBankName.BoundText = ""
        'TxtChequeNumber.text = ""

        Frame3.Enabled = True
        Frame4.Visible = True
    ElseIf Me.CboPayMentType.ListIndex = 2 Then
    txtTransferExpenses.Enabled = True
    Cmd(13).Enabled = True
        Frame11.Visible = True
        Frame14.Visible = True
        Frame15.Visible = True
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        '    DcboBankName.BoundText = ""
        'TxtChequeNumber.text = ""

        Frame3.Enabled = True
        Frame4.Visible = True

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(16).Caption = "—ﬁ„ «·ÕÊ«·… "
            lbl(17).Caption = " «—ÌŒÂ«"
    
        Else
            lbl(16).Caption = "Transfer No"
            lbl(17).Caption = "Date"
        End If
  
    Else
        Frame11.Visible = True
        Frame14.Visible = True
        Frame15.Visible = True
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Cmd(13).Enabled = True
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub Check18_Click()
    Dim i As Integer

    If Check18.value = vbChecked Then

        With Me.GRID1
 
            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.GRID1

            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If
    Reline
End Sub


Public Sub Reline2()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.GRID1
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And .cell(flexcpChecked, i, .ColIndex("haveqest")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                .TextMatrix(i, .ColIndex("TransPayedValue")) = val(.TextMatrix(i, .ColIndex("InstalValue")))
           Sm = Sm + val(.TextMatrix(i, .ColIndex("InstalValue")))
           End If
           Next i
  
    End With
    If Sm > 0 Then
    XPTxtVal.text = Sm
   XPTxtVal.Enabled = False
   CalCuteCurrency
   Else
  ' XPTxtVal.Enabled = True
   End If
End Sub
Sub RelineProject22()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid2
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
              If val(.TextMatrix(i, .ColIndex("TransPayedValue"))) > val(.TextMatrix(i, .ColIndex("RemainingValue"))) And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) <> 0 Then
              If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ﬁÌ„… «·œ›⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
              Else
              MsgBox "Can Not PaymentValue Larger Than Total Value "
              End If
              .TextMatrix(i, .ColIndex("TransPayedValue")) = 0
              Exit Sub
              End If
           Sm = Sm + val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           End If
           Next i
  
    End With
   XPTxtVal.text = Sm
   XPTxtVal.Enabled = False
   CalCuteCurrency
End Sub
Sub RelineBu22()
    Dim IntCounter As Integer
    Dim Sm As Double
    Dim Sm2 As Double
    Sm = 0
    Sm2 = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid1
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                
              If val(.TextMatrix(i, .ColIndex("TransPayedValue"))) > val(.TextMatrix(i, .ColIndex("RemainingValue"))) And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) <> 0 Then
              If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ﬁÌ„… «·œ›⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
              Else
              MsgBox "Can Not PaymentValue Larger Than Total Value "
              End If
              .TextMatrix(i, .ColIndex("TransPayedValue")) = 0
              Exit Sub
              End If
              
           Sm = Sm + val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           End If
           Next i
  
    End With
   XPTxtVal.text = Sm
   XPTxtVal.Enabled = False
   CalCuteCurrency
End Sub
Sub Reline22()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    If Aut_manual = True Then
    Dim i As Integer
    With Me.GRID1
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
              If val(.TextMatrix(i, .ColIndex("TransPayedValue"))) > val(.TextMatrix(i, .ColIndex("RemainingValue"))) And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) <> 0 Then
              If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ﬁÌ„… «·œ›⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
              Else
              MsgBox "Can Not PaymentValue Larger Than Total Value "
              End If
              .TextMatrix(i, .ColIndex("TransPayedValue")) = 0
              Exit Sub
              End If
           Sm = Sm + val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           End If
           Next i
  
    End With
   XPTxtVal.text = Sm
   XPTxtVal.Enabled = False
   Else
   XPTxtVal.Enabled = True
   End If
   CalCuteCurrency
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
  Label27(1).Caption = Sm
End Sub
Sub RelineProject()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid2
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           Sm = Sm + val(.TextMatrix(i, .ColIndex("RemainingValue")))
           End If
           Next i
  
    End With
  Label27(3).Caption = Sm
End Sub
Sub Reline()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.GRID1
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           Sm = Sm + val(.TextMatrix(i, .ColIndex("RemainingValue")))
           End If
           Next i
  
    End With
   Label16.Caption = Sm
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

Function sand_numbering() As String
    

     
End Function

Function GetComponentValuePerBranch4(BramchId As Integer, componentname As String) As Double
    Dim SUM As Double
    SUM = 0
    Dim i As Integer




If DCboCashType.ListIndex <> 9 Then Exit Function

    With FrmEmpSalary6.VSFlexGrid2

        For i = .FixedRows To .rows - 1
    
            If .cell(flexcpChecked, i, .ColIndex("Ch")) = flexChecked And val(.TextMatrix(i, .ColIndex("net"))) > 0 And val(.TextMatrix(i, .ColIndex("BranchId"))) = BramchId Then
                SUM = SUM + val(.TextMatrix(i, .ColIndex("net")))
            End If

        Next i

    End With

    GetComponentValuePerBranch4 = SUM
End Function


Function GetComponentValuePerBranch3(BramchId As Integer, componentname As String) As Double
    Dim SUM As Double
    SUM = 0
    Dim i As Integer

If DCboCashType.ListIndex <> 7 Then Exit Function

    With FrmEmpSalary6.VSFlexGrid1

        For i = .FixedRows To .rows - 2
    
            If .cell(flexcpChecked, i, .ColIndex("Ch")) = flexChecked And val(.TextMatrix(i, .ColIndex("Valu"))) > 0 And val(.TextMatrix(i, .ColIndex("BranchId"))) = BramchId Then
                SUM = SUM + val(.TextMatrix(i, .ColIndex("Valu")))
            End If

        Next i

    End With

    GetComponentValuePerBranch3 = SUM
End Function
Function GetComponentValuePerBranch16(BramchId As Integer, componentname As String) As Double
    Dim SUM As Double
    SUM = 0
    Dim i As Integer

If DCboCashType.ListIndex <> 11 Then Exit Function

    With FrmEmpSalary6.VSFlexGrid4

        For i = .FixedRows To .rows - 2
    
            If .cell(flexcpChecked, i, .ColIndex("Ch")) = flexChecked And val(.TextMatrix(i, .ColIndex("MordValue"))) > 0 And val(.TextMatrix(i, .ColIndex("BrnchID1"))) = BramchId Then
                SUM = SUM + val(.TextMatrix(i, .ColIndex("MordValue")))
            End If

        Next i

    End With

    GetComponentValuePerBranch16 = SUM
End Function

Function GetComponentValuePerBranch2(BramchId As Integer, componentname As String) As Double
    Dim SUM As Double
    SUM = 0
    Dim i As Integer
If DCboCashType.ListIndex <> 6 Then Exit Function
    With FrmEmpSalary6.GRID1

        For i = .FixedRows To .rows - 1
    
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And val(.TextMatrix(i, .ColIndex(componentname))) > 0 And val(.TextMatrix(i, .ColIndex("BranchId"))) = BramchId Then
                SUM = SUM + val(.TextMatrix(i, .ColIndex(componentname)))
            End If

        Next i

    End With

    GetComponentValuePerBranch2 = SUM
End Function
Function payGl1Suppler(LngDevID As Long, notes_id As Double) As Double
If DCboCashType.ListIndex <> 9 Then Exit Function
Dim rsBranch As New ADODB.Recordset
Dim total_value As Double
Dim cProgress As ClsProgress
 Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double
     My_SQL = "SELECT  (branch_id) From TblBranchesData"
   
    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsBranch.RecordCount > 0 Then
        rsBranch.MoveFirst
    End If
    
    
    cProgress.StartProgress

    DoEvents
    total_value = XPTxtVal.text
    
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    'total_value = 0
                
    Dim BranchID As Integer
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
    Dim StrSQL As String
Dim i As Integer
Dim Msg2 As String
Dim Msgn As String
Msg = " ”‰œ ’—› ··„ ⁄ÂœÌ‰  " & CHR(13) & "  ·ÿ·» «·’—› —ﬁ„ " & TxtOrderSuppler & CHR(13) & " ··„‰ÿﬁ…  " & DcbBrReq.text & CHR(13) & " ··› —…  " & dcDur.text & CHR(13) & "  ·‘Â— " & dcMontth.text
'Msgn = " ”‰œ ’—› ··„ ⁄ÂœÌ‰ »—ﬁ„  " & TxtNoteSerial1.text & Chr(13) & "  ·ÿ·» «·’—› —ﬁ„ " & TxtOrderSuppler & Chr(13) & " ··„‰ÿﬁ…  " & DcbBrReq.text & Chr(13) & " ··› —…  " & DcDur.text & Chr(13) & "  ·‘Â— " & dcMontth.text

    With FrmEmpSalary6.VSFlexGrid2
OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
        For i = .FixedRows To .rows - 1
 
            If .cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
          
           
            
                BranchID = val(.TextMatrix(i, .ColIndex("BranchId")))
               total_value = total_value + Round(val(.TextMatrix(i, .ColIndex("net"))), 2)
            
                depit_side = .TextMatrix(i, .ColIndex("Account_Code"))
                
                CURRENT_LINE = setfoxy_Line

                If val(.TextMatrix(i, .ColIndex("net"))) > 0 Then
                
         


       Msg2 = CHR(13) & "  —ﬁ„ «·„⁄œÂ/«·”Ì«—…" & .TextMatrix(i, .ColIndex("BoardNO"))
                    If ModAccounts.AddNewDev(LngDevID, i + 1, depit_side, Round(.TextMatrix(i, .ColIndex("net")), 2), 0, Msg + " " + Msg2, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        'GoTo ErrTrap
                    End If
                End If
                 StrSQL = "Update TblAttributionInstallmentDivided Set  PayMentPayed=1 ,noteID=" & notes_id & ",noteserial1='" & TxtNoteSerial1 & "' Where ID=" & val(.TextMatrix(i, .ColIndex("InsID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
              
                         End If
                     

        Next i

    End With
    
    
          Dim Branch As Integer
            Dim CValue  As Double
Dim DeptSide1 As String
Dim credit_side1 As String

            If rsBranch.RecordCount > 0 Then
                rsBranch.MoveFirst
            End If

            i = i + 1

                                    For Branch = 1 To rsBranch.RecordCount
                                                                                                 
                                        BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                                                    
                                                    
                                        CValue = GetComponentValuePerBranch4(BranchID, "")
                                                   If BranchID = val(Me.dcBranch.BoundText) Then CValue = 0
                                                   
                                                 DeptSide1 = getBranchCurrentAccount(BranchID)
                                                 credit_side1 = getBranchCurrentAccount(dcBranch.BoundText)
                                                    If CValue > 0 Then
                                                                                        
                                                                    If CValue > 0 Then
                                                                    OtherInformation.NextAccount_Code = credit_side1
                                                                                    If ModAccounts.AddNewDev(LngDevID, i, DeptSide1, CValue, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                                                    OtherInformation.NextAccount_Code = DeptSide1
                                                         i = i + 1
                                                                If ModAccounts.AddNewDev(LngDevID, i, credit_side1, CValue, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                                                    
                                                                        i = i + 1
                                                                    End If
                                                                                                                
                                                    End If
                        
                                        rsBranch.MoveNext
                                    Next Branch

         
                
     
 payGl1Suppler = i + 1
     
     
ErrTrap:
End Function
Function payGl16(LngDevID As Long, notes_id As Double) As Double
If DCboCashType.ListIndex <> 11 Then Exit Function
Dim rsBranch As New ADODB.Recordset
Dim total_value As Double
Dim cProgress As ClsProgress
 Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double
     My_SQL = "SELECT  (branch_id) From TblBranchesData"
   OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsBranch.RecordCount > 0 Then
        rsBranch.MoveFirst
    End If
    
    
    cProgress.StartProgress

    DoEvents
    total_value = XPTxtVal.text
    
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    'total_value = 0
                
    Dim BranchID As Integer
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
    Dim StrSQL As String
Dim i As Integer
    With FrmEmpSalary6.VSFlexGrid4
        For i = .FixedRows To .rows - 2
            If .cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
                BranchID = val(.TextMatrix(i, .ColIndex("BrnchID1")))
               total_value = total_value + Round(val(.TextMatrix(i, .ColIndex("MordValue"))), 2)
            
                depit_side = .TextMatrix(i, .ColIndex("Account_Code"))
                
                CURRENT_LINE = setfoxy_Line

                If val(.TextMatrix(i, .ColIndex("MordValue"))) > 0 Then
                
                
                    If ModAccounts.AddNewDev(LngDevID, i + 1, depit_side, Round(.TextMatrix(i, .ColIndex("MordValue")), 2), 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(1), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        'GoTo ErrTrap
                    End If
                    StrSQL = "Update TblApproveCompoYearDet Set  PaymentPayed=1  Where ID=" & val(.TextMatrix(i, .ColIndex("ID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                 StrSQL = "Update TblComponentYearDet Set  PaymentPayed=1  Where ID=" & val(.TextMatrix(i, .ColIndex("CompYerID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
              
                End If
               
              
                         End If
                     

        Next i

    End With
    
    
          Dim Branch As Integer
            Dim CValue  As Double
Dim DeptSide1 As String
Dim credit_side1 As String

            If rsBranch.RecordCount > 0 Then
                rsBranch.MoveFirst
            End If

            i = i + 1

                                    For Branch = 1 To rsBranch.RecordCount
                                                                                                 
                                        BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                                                    
                                                    
                                        CValue = GetComponentValuePerBranch16(BranchID, "EmpTotalNet")
                                                   If BranchID = val(Me.dcBranch.BoundText) Then CValue = 0
                                                   
                                                 DeptSide1 = getBranchCurrentAccount(BranchID)
                                                 credit_side1 = getBranchCurrentAccount(dcBranch.BoundText)
                                                    If CValue > 0 Then
                                                                                        
                                                                    If CValue > 0 Then
                                                                    OtherInformation.NextAccount_Code = credit_side1
                                                                                    If ModAccounts.AddNewDev(LngDevID, i, DeptSide1, CValue, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                         i = i + 1
                                                         OtherInformation.NextAccount_Code = DeptSide1
                                                                If ModAccounts.AddNewDev(LngDevID, i, credit_side1, CValue, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                                                    
                                                                        i = i + 1
                                                                    End If
                                                                                                                
                                                    End If
                        
                                        rsBranch.MoveNext
                                    Next Branch

         
                
     
 payGl16 = i + 1
     
ErrTrap:
End Function
Function payGl122(LngDevID As Long, notes_id As Double) As Double
If DCboCashType.ListIndex <> 7 Then Exit Function
Dim rsBranch As New ADODB.Recordset
Dim total_value As Double
Dim cProgress As ClsProgress
 Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double
     My_SQL = "SELECT  (branch_id) From TblBranchesData"
   
    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsBranch.RecordCount > 0 Then
        rsBranch.MoveFirst
    End If
    
    
    cProgress.StartProgress

    DoEvents
    total_value = XPTxtVal.text
    
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    'total_value = 0
                
    Dim BranchID As Integer
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
    Dim StrSQL As String
Dim i As Integer
    With FrmEmpSalary6.VSFlexGrid1
OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
        For i = .FixedRows To .rows - 2
 
            If .cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
          
           
            Msg = XPMTxtRemarks & " - " & txt_general_des.text
                BranchID = val(.TextMatrix(i, .ColIndex("BranchId")))
               total_value = total_value + Round(val(.TextMatrix(i, .ColIndex("Valu"))), 2)
            
                depit_side = .TextMatrix(i, .ColIndex("Account_Code"))
                DcboDebitSide.BoundText = depit_side
                CURRENT_LINE = setfoxy_Line

                If val(.TextMatrix(i, .ColIndex("Valu"))) > 0 Then
                
                
                    If ModAccounts.AddNewDev(LngDevID, i + 1, depit_side, Round(.TextMatrix(i, .ColIndex("Valu")), 2), 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(1), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        'GoTo ErrTrap
                    End If
                    StrSQL = "Update TblPripaidExpensesDet Set  PaymentPayed=1  Where ID=" & val(.TextMatrix(i, .ColIndex("MainID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
              
                End If
               
              
                         End If
                     

        Next i

    End With
    
    
          Dim Branch As Integer
            Dim CValue  As Double
Dim DeptSide1 As String
Dim credit_side1 As String

            If rsBranch.RecordCount > 0 Then
                rsBranch.MoveFirst
            End If

            i = i + 1

                                    For Branch = 1 To rsBranch.RecordCount
                                                                                                 
                                        BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                                                    
                                                    
                                        CValue = GetComponentValuePerBranch3(BranchID, "EmpTotalNet")
                                                   If BranchID = val(Me.dcBranch.BoundText) Then CValue = 0
                                                   
                                                 DeptSide1 = getBranchCurrentAccount(BranchID)
                                                 credit_side1 = getBranchCurrentAccount(dcBranch.BoundText)
                                                    If CValue > 0 Then
                                                                                        
                                                                    If CValue > 0 Then
                                                                                    If ModAccounts.AddNewDev(LngDevID, i, DeptSide1, CValue, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                         i = i + 1
                                                                If ModAccounts.AddNewDev(LngDevID, i, credit_side1, CValue, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                                                    
                                                                        i = i + 1
                                                                    End If
                                                                                                                
                                                    End If
                        
                                        rsBranch.MoveNext
                                    Next Branch
  i = i + 1
                              If val(TxtPrePayd(17).text) > 0 Then
                              Dim AccountVATCreit As String
'                              GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
'                                       If ModAccounts.AddNewDev(LngDevID, i, AccountVATCreit, val(TxtPrePayd(17).Text), 0, Msg & "Õ”«» «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                                          GoTo ErrTrap
'                                       End If
                             End If
                
     
 payGl122 = i + 1
     
     
ErrTrap:
End Function
'Function GetComponentValuePerBranch2(BramchId As Integer, componentname As String) As Double
'    Dim SUM As Double
'    SUM = 0
  '  Dim i As Integer
'
'    With FrmEmpSalary6.Grid1
'
'        For i = .FixedRows To .Rows - 2
'
'            If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And val(.TextMatrix(i, .ColIndex(componentname))) > 0 And val(.TextMatrix(i, .ColIndex("BranchId"))) = BramchId Then
'                SUM = SUM + val(.TextMatrix(i, .ColIndex(componentname)))
'            End If
'
'        Next i

'    End With

'    GetComponentValuePerBranch2 = SUM
'End Function
'

Function payGl10(LngDevID As Long, notes_id As Double) As Double

If DCboCashType.ListIndex <> 10 Then Exit Function
Dim rsBranch As New ADODB.Recordset
Dim total_value As Double
Dim cProgress As ClsProgress
 Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double
 
    DoEvents
    total_value = XPTxtVal.text
    
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    'total_value = 0
                
    Dim BranchID As Integer
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
Dim i As Integer
Dim lineno As Integer
If SystemOptions.UserInterface = ArabicInterface Then
Msg = " „œ›Ê⁄«  ‰Â«Ì… «·Œœ„… —ﬁ„ " + TxtEndService.text + "”‰œ œ›⁄ —ﬁ„" + TxtNoteSerial1
Else
Msg = "Payments End Service No " + TxtEndService.text + "No" + TxtNoteSerial1

End If
lineno = 1
''******************************************«ÃÊ— „” Õﬁ…
                 BranchID = DcbBranchEndServ.BoundText
                total_value = val(txtSal.text)
    OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
                depit_side = get_EMPLOYEE_Account(DcbEmpEndService.BoundText, "Account_Code1")
                CURRENT_LINE = setfoxy_Line

                If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " —« » ‘Â— Õ«·Ì  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
                      ''/////////////‰Â«Ì… «·Œœ„…
              depit_side = get_EMPLOYEE_Account(DcbEmpEndService.BoundText, "Account_Code4")
                total_value = val(txtTotal.text)   'Endsev
                BranchID = DcbBranchEndServ.BoundText
               If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " ﬁÌ„… ‰Â«Ì… «·Œœ„…  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
                            ''////////////////// «Ã«“…
                            depit_side = get_account_code_branch(141, my_branch)
                   ' depit_side = get_EMPLOYEE_Account(DcbEmpEndService.BoundText, "Account_Code2")
               ' total_value = val(txtCustom.text)   'Endsev
               total_value = val(TxtCash.text)
                BranchID = DcbBranchEndServ.BoundText
               If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + "  «·«Ã«“… »œÊ‰ —« » ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
                          ''//////////////////„Œ’’  –«ﬂ—
                    depit_side = get_EMPLOYEE_Account(DcbEmpEndService.BoundText, "Account_Code5")
                total_value = val(txtTicketValue.text)   'Endsev
                BranchID = DcbBranchEndServ.BoundText
               If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + "  –«ﬂ—  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
                
                
                                          ''//////////////////„Œ’’ «Ã«“…
                    depit_side = get_EMPLOYEE_Account(DcbEmpEndService.BoundText, "Account_Code2")
                total_value = val(txtCustom2.text)   'Endsev
                BranchID = DcbBranchEndServ.BoundText
               If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " „Œ’’ «Ã«“Â  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
                
                
              '''/////////////////////////////////
          ''///////////////////«÷«›«  ‰Â«Ì… «·Œœ„…
                 total_value = val(TxtAddOther2.text)
                 my_branch = BranchID
               ' depit_side = get_account_code_branch(139, my_branch)
                depit_side = get_EMPLOYEE_Account(DcbEmpEndService.BoundText, "Account_Code4")
                CURRENT_LINE = setfoxy_Line

                If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " «÷«›Ì ‰Â«Ì… «·Œœ„…  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                    
                End If
          '''//////////////////Œ’Ê„«  —Ê« »
                   total_value = val(TxtCash.text)
                 my_branch = BranchID
             '   depit_side = get_account_code_branch(139, my_branch)
               depit_side = get_EMPLOYEE_Account(DcbEmpEndService.BoundText, "Account_Code")
                CURRENT_LINE = setfoxy_Line

                If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + "  ‰Â«Ì… «·Œœ„… «Ã«“«  »œÊ‰ —« » ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If

                    ''///////////////////Œ’Ê„« Œ’Ê„«  ‰Â«Ì… «·Œœ„…
                 total_value = val(TxtPrePayd(13).text)
                 my_branch = BranchID
             '  depit_side = get_account_code_branch(139, my_branch)
                depit_side = get_EMPLOYEE_Account(DcbEmpEndService.BoundText, "Account_Code")
                CURRENT_LINE = setfoxy_Line

                If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + "  ‰Â«Ì… «·Œœ„… Œ’Ê„«  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
                                   ''///////////////////«Ã«“«  »œÊ‰ —« » ‰Â«Ì… «·Œœ„…
                 total_value = val(TxtPrePayd(16).text)
                 my_branch = BranchID
            ' depit_side = get_account_code_branch(139, my_branch)
              depit_side = get_EMPLOYEE_Account(DcbEmpEndService.BoundText, "Account_Code")
                CURRENT_LINE = setfoxy_Line

                If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + "  ‰Â«Ì… «·Œœ„… «Ã«“«  »œÊ‰ —« » ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
       
       
            ''//////////////////////”·›
                    depit_side = get_EMPLOYEE_Account(DcbEmpEndService.BoundText, "Account_code")
                total_value = val(TXTAdvanceTotal.text)   'Endsev
                BranchID = val(DcbBranchEndServ.BoundText)
               If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + " ”·›…  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
                
            ''/////////////////////«Ã«“«  »œÊ‰ —« »
            '  total_value = val(TxtCash.text)
            '     my_branch = BranchID
            '    depit_side = get_account_code_branch(141, my_branch)
            '       If total_value > 0 Then
            '        If ModAccounts.AddNewDev(LngDevID, LineNo, depit_side, Round(total_value, 2), 1, Msg + " «Ã«“«  »œÊ‰ —« »  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText)) = False Then
            '            GoTo ErrTrap
            '        End If
            '        LineNo = LineNo + 1
            '    End If
                ''////////////

Dim sql As String
Dim Account_code2 As String
Dim Balance As String
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
sql = sql & " SELECT     Account_Code, empid,Type"
sql = sql & " From dbo.TblBoxesData"
sql = sql & " Where (EmpID = " & val(DcbEmpEndService.BoundText) & ") and Type=1"
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
Rs9.MoveFirst
For i = 1 To Rs9.RecordCount
depit_side = IIf(IsNull(Rs9("Account_Code").value), "", Rs9("Account_Code").value)
WriteCustomerBalPublic depit_side, Balance
                total_value = val(Balance)
                BranchID = DcbBranchEndServ.BoundText
               If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + " ⁄Âœ…  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcbEmpEndService.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                End If
Rs9.MoveNext
Next i
End If
                
 
 
    If val(XPTxtVal) > 0 Then
                        
             Dim Branch As Integer
            Dim CValue  As Double
Dim DeptSide1 As String
Dim credit_side1 As String
                                                                                                  
                                        BranchID = val(DcbBranchEndServ.BoundText)
                                                    
                                                    
                                        CValue = XPTxtVal
                                                   If BranchID = val(Me.dcBranch.BoundText) Then CValue = 0
                                                   
                                                 DeptSide1 = getBranchCurrentAccount(BranchID)
                                                 credit_side1 = getBranchCurrentAccount(dcBranch.BoundText)
                                                    If CValue > 0 Then
                                                                                        
                                                                    If CValue > 0 Then
                                                                    OtherInformation.NextAccount_Code = credit_side1
                                                                                    If ModAccounts.AddNewDev(LngDevID, lineno, DeptSide1, CValue, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                                                    OtherInformation.NextAccount_Code = DeptSide1
                                                         lineno = lineno + 1
                                                                If ModAccounts.AddNewDev(LngDevID, lineno, credit_side1, CValue, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                                                    
                                                                        lineno = lineno + 1
                                                                    End If
                                                                                                                
                                                    End If
                        
                               
        End If
                 
 payGl10 = lineno + 1
      Dim X As Integer
   
     

    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
ErrTrap:
 
End Function

Function payGl8(LngDevID As Long, notes_id As Double) As Double

If DCboCashType.ListIndex <> 8 Then Exit Function
Dim rsBranch As New ADODB.Recordset
Dim total_value As Double
Dim cProgress As ClsProgress
 Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double
    DoEvents
    total_value = XPTxtVal.text
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    'total_value = 0
     OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
    Dim BranchID As Integer
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
Dim i As Integer
Dim lineno As Integer
Msg = " „œ›Ê⁄«  „” Õﬁ«  «Ã«“… —ﬁ„ " + TxtDue + "”‰œ œ›⁄ —ﬁ„" + TxtNoteSerial1
lineno = 1
''******************************************«ÃÊ— „” Õﬁ…
                 BranchID = val(dcBranch1.BoundText)
                 If val(TxtSalary.text) > val(TxtInsuranceValue) Then
                total_value = val(TxtSalary) - val(TxtInsuranceValue)
              Else
              total_value = val(TxtSalary)
              End If
                    Dim insurancedes As String
            If val(TxtInsuranceValue) > 0 And val(TxtSalary.text) > val(TxtInsuranceValue) Then
             insurancedes = CHR(13) & " „ Œ’„ «· √„Ì‰ ·› —… «·«Ã«“… »ﬁÌ„Â " & val(TxtInsuranceValue.text)
           
            End If
                depit_side = get_EMPLOYEE_Account(val(DcboEmpName.BoundText), "Account_Code1")
                CURRENT_LINE = setfoxy_Line

                If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " —« » ‘Â— Õ«·Ì  " + insurancedes, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcboEmpName.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                ElseIf total_value < 0 Then
                    GoTo ErrTrap
                End If
              
                total_value = val(TxtSalEntitOther) 'addition
               If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " —« » ‘Â— Õ«·Ì  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcboEmpName.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                ElseIf total_value < 0 Then
                    GoTo ErrTrap
                End If
                
                
           total_value = val(Txtother) 'Discounts
               If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + " —« » ‘Â— Õ«·Ì  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcboEmpName.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                ElseIf total_value < 0 Then
                    GoTo ErrTrap
                End If
                
                
            total_value = val(txtAdvance1) 'Advance
               If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 1, Msg + " —« » ‘Â— Õ«·Ì  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcboEmpName.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                ElseIf total_value < 0 Then
                    GoTo ErrTrap
                End If
                
                
 
''******************************************«Ã«“…
                 BranchID = val(dcBranch1.BoundText)
            If val(TxtInsuranceValue) > 0 And val(TxtSalary.text) < val(TxtInsuranceValue) Then
             insurancedes = CHR(13) & " „ Œ’„ «· √„Ì‰ ·› —… «·«Ã«“… »ﬁÌ„Â " & val(TxtInsuranceValue.text)
           
            End If
                 If val(TxtSalary) < val(TxtInsuranceValue) Then
                 total_value = val(txtSalaryVocation.text) - val(TxtInsuranceValue)
                 Else
                total_value = val(txtSalaryVocation.text)
                End If
            
                depit_side = get_EMPLOYEE_Account(DcboEmpName.BoundText, "Account_Code2")
                CURRENT_LINE = setfoxy_Line

                If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + "   „” Õﬁ«  «Ã«“…" + insurancedes, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcboEmpName.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                ElseIf total_value < 0 Then
                        GoTo ErrTrap
                End If
 ''****************************************** –«ﬂ—
                 BranchID = val(dcBranch1.BoundText)
                total_value = val(Me.txtValueTickt)
            
                depit_side = get_EMPLOYEE_Account(DcboEmpName.BoundText, "Account_Code5")
                CURRENT_LINE = setfoxy_Line

                If val(txtValueTickt) > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, depit_side, Round(total_value, 2), 0, Msg + " «ÃÊ—  –«ﬂ—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(DcboEmpName.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                ElseIf val(txtValueTickt) < 0 Then
                    GoTo ErrTrap
                End If
    If val(XPTxtVal) > 0 Then
                        
             Dim Branch As Integer
            Dim CValue  As Double
Dim DeptSide1 As String
Dim credit_side1 As String
                                                                                                  
                                        BranchID = val(dcBranch1.BoundText)
                                                    
                                                    
                                        CValue = val(XPTxtVal.text)
                                                   If BranchID = val(Me.dcBranch.BoundText) Then CValue = 0
                                                   
                                                 DeptSide1 = getBranchCurrentAccount(BranchID)
                                                 credit_side1 = getBranchCurrentAccount(dcBranch.BoundText)
                                                    If CValue > 0 Then
                                                                                        
                                                                    If CValue > 0 Then
                                                                    OtherInformation.NextAccount_Code = credit_side1
                                                                    
                                                                                    If ModAccounts.AddNewDev(LngDevID, lineno, DeptSide1, CValue, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                                                    OtherInformation.NextAccount_Code = DeptSide1
                                                         lineno = lineno + 1
                                                                If ModAccounts.AddNewDev(LngDevID, lineno, credit_side1, CValue, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                                                    
                                                                        lineno = lineno + 1
                                                                    End If
                                                                                                                
                                                    End If
                        
                               
        End If
                 
 payGl8 = lineno + 1
      Dim X As Integer
   
     

    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
    Exit Function
ErrTrap:
    payGl8 = -1
        
       '     LogTextA = "    ‘«‘…  „”Ì— «·—Ê« »   „ «‰‘«¡ «·ﬁÌœ ··—Ê« » Ê«·„”Ì— " & Chr(13) & " «·‘Â—     " & CmbMonth.text & Chr(13) & "  «·”‰…   " & CboYear.text & Chr(13) & " «· «—ÌŒ " & DTP_Date.value
                     
   ' LogTextE = ""
   '    AddToLogFile CInt(user_id), 555, Date, Time, LogTextA, LogTextE, Me.name, "N", "", , val(TxtNoteSerial), ""
 
 
End Function
Function payGl1(LngDevID As Long, notes_id As Double) As Double

  
 If DCboCashType.ListIndex <> 1 Then Exit Function
Dim rsBranch As New ADODB.Recordset
Dim total_value As Double
Dim total_valuee As Double
Dim cProgress As ClsProgress
 Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Integer
 
 
    DoEvents
    total_value = XPTxtVal.text
    
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    'total_value = 0
                Line1 = 2
    Dim BranchID As Integer
    Dim BranchID2 As Integer
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
    Dim DeptSide As String
    Dim DeptSide1 As String
    Dim CreditSide1 As String
    Dim StrSQL As String
    Dim k As Integer
    k = 0
Dim i As Integer
Line1 = 3
    With GRID1

        For i = .FixedRows To .rows - 1
 
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) > 0 Then
          
           BranchID = val(Me.dcBranch.BoundText)
            
                BranchID2 = val(.TextMatrix(i, .ColIndex("branch_no")))
                
                                  DeptSide = getBranchCurrentAccount(BranchID)
                                                 credit_side = getBranchCurrentAccount(BranchID2)
                                                 DeptSide1 = DcboDebitSide.BoundText
                                                 CreditSide1 = DcboCreditSide.BoundText
                                                 
                
            
                total_value = Round(.TextMatrix(i, .ColIndex("TransPayedValue")), 2)
              
               If val(TxtCurrencyRate.text) = 0 Then
               TxtCurrencyRate.text = 1
               End If
                total_valuee = Round(total_valuee / val(TxtCurrencyRate.text), 2)
 '  Dim DeptSide As String
  '  Dim CreditSide As String
             
                CURRENT_LINE = setfoxy_Line

                If total_value > 0 Then
                
             Msg = "  ”œ«œ Ã“¡ „‰ ›« Ê—… „«·Ì…" & CHR(13) & .TextMatrix(i, .ColIndex("NoteSerial1"))
              Msg = Msg & CHR(13) & "··›—⁄ " & .TextMatrix(i, .ColIndex("branch_name"))
              Msg = Msg & CHR(13) & "”œœ  „‰  " & dcBranch.text
                          
                                      '„Ê—œ
                                      OtherInformation.NextAccount_Code = CreditSide1
                                        If ModAccounts.AddNewDev(LngDevID, Line1, DeptSide1, total_value, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , total_valuee, Me.DcbCurrency.text, TxtCurrencyRate.text, , , CURRENT_LINE, , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              
                                                              Line1 = Line1 + 1
                                                               OtherInformation.NextAccount_Code = DeptSide1
                                                  '⁄Âœ…
                                                  If ModAccounts.AddNewDev(LngDevID, Line1, CreditSide1, total_value, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , total_valuee, DcbCurrency.text, TxtCurrencyRate.text, , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              
                                                                      Line1 = Line1 + 1
                                          If BranchID <> BranchID2 Then
                                                              Line1 = Line1 + 1
                                                  'Ã —Ì
                                                  OtherInformation.NextAccount_Code = DeptSide
                                               If ModAccounts.AddNewDev(LngDevID, Line1, credit_side, total_value, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , total_valuee, DcbCurrency.text, TxtCurrencyRate.text, , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              
                                                              Line1 = Line1 + 1
                                                              OtherInformation.NextAccount_Code = credit_side
                                                        'Ã«—Ì
                                                              If ModAccounts.AddNewDev(LngDevID, Line1, DeptSide, total_value, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , total_valuee, DcbCurrency.text, TxtCurrencyRate.text, , , CURRENT_LINE, , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              
                                                        
                                    
                                                        
                                        Line1 = Line1 + 1
                                        End If
               
                                
            End If
                             
                     
End If
        Next i
     ' total_value = Round(val(txtTransferExpenses.Text), 2)
     ' total_valuee = total_value / 1
     'If total_value > 0 Then
' If ModAccounts.AddNewDev(LngDevID, Line1, DcboCreditSide.BoundText, total_value, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , total_valuee, DcbCurrency.Text, TxtCurrencyRate.Text, , , CURRENT_LINE, , , , , , , , , BranchID) = False Then
'
'  End If
' End If
    End With
 payGl1 = Line1 + 1
     
     
ErrTrap:
End Function
Function payGlBillBuy1(LngDevID As Long, notes_id As Double) As Double

  
 If DCboCashType.ListIndex <> 1 Then Exit Function
Dim rsBranch As New ADODB.Recordset
Dim total_value As Double
Dim total_valuee As Double
Dim cProgress As ClsProgress
 Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Integer
 
 
    DoEvents
    OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
    total_value = XPTxtVal.text
    
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    'total_value = 0
                Line1 = 2
    Dim BranchID As Integer
    Dim BranchID2 As Integer
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
    Dim DeptSide As String
    Dim DeptSide1 As String
    Dim CreditSide1 As String
    Dim StrSQL As String
    Dim k As Integer
    k = 0
Dim i As Integer
Line1 = 3
    With VSFlexGrid1

        For i = .FixedRows To .rows - 1
 
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) > 0 Then
          
           BranchID = val(Me.dcBranch.BoundText)
            
                BranchID2 = val(.TextMatrix(i, .ColIndex("branch_no")))
                
                                  DeptSide = getBranchCurrentAccount(BranchID)
                                                 credit_side = getBranchCurrentAccount(BranchID2)
                                                 DeptSide1 = DcboDebitSide.BoundText
                                                 CreditSide1 = DcboCreditSide.BoundText
                                                 
              ' If k = 0 Then
              '  total_value = Round(.TextMatrix(I, .ColIndex("TransPayedValue")) + val(txtTransferExpenses.Text), 2)
              '  k = 1
              ' Else
                total_value = Round(.TextMatrix(i, .ColIndex("TransPayedValue")), 2)
              ' End If
                If val(TxtCurrencyRate.text) = 0 Then
                TxtCurrencyRate.text = 1
                End If
                total_valuee = Round(total_value / val(TxtCurrencyRate.text), 2)
 '  Dim DeptSide As String
  '  Dim CreditSide As String
             
                CURRENT_LINE = setfoxy_Line

                If total_value > 0 Then
                
             Msg = "  ”œ«œ Ã“¡ „‰ ›« Ê—… „‘ —Ì« " & CHR(13) & .TextMatrix(i, .ColIndex("NoteSerial1"))
              Msg = Msg & CHR(13) & "··›—⁄ " & .TextMatrix(i, .ColIndex("branch_name"))
              Msg = Msg & CHR(13) & "”œœ  „‰  " & dcBranch.text
                          
                                      '„Ê—œ
                                      OtherInformation.NextAccount_Code = CreditSide1
                                        If ModAccounts.AddNewDev(LngDevID, Line1, DeptSide1, total_value, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , total_valuee, Me.DcbCurrency.text, TxtCurrencyRate.text, , , CURRENT_LINE, , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              OtherInformation.NextAccount_Code = DeptSide1
                                                              Line1 = Line1 + 1
                                                  '⁄Âœ…
                                                  
                                                  If ModAccounts.AddNewDev(LngDevID, Line1, CreditSide1, total_value, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , total_valuee, DcbCurrency.text, TxtCurrencyRate.text, , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              
                                                                      Line1 = Line1 + 1
                                          If BranchID <> BranchID2 Then
                                                              Line1 = Line1 + 1
                                                  'Ã —Ì
                                                  OtherInformation.NextAccount_Code = DeptSide
                                               If ModAccounts.AddNewDev(LngDevID, Line1, credit_side, total_value, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , total_valuee, DcbCurrency.text, TxtCurrencyRate.text, , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              OtherInformation.NextAccount_Code = credit_side
                                                              Line1 = Line1 + 1
                                                        'Ã«—Ì
                                                              If ModAccounts.AddNewDev(LngDevID, Line1, DeptSide, total_value, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , total_valuee, DcbCurrency.text, TxtCurrencyRate.text, , , CURRENT_LINE, , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              
                                                        
                                    
                                                        
                                        Line1 = Line1 + 1
                                        End If
               
                                
            End If
                             
                     
End If
        Next i
   '   total_value = Round(val(txtTransferExpenses.Text), 2)
   '   total_valuee = total_value / 1
   '  If total_value > 0 Then
 'If ModAccounts.AddNewDev(LngDevID, Line1, CreditSide1, total_value, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , total_valuee, DcbCurrency.Text, TxtCurrencyRate.Text, , , CURRENT_LINE, , , , , , , , , BranchID) = False Then
                                                                   
 ' End If
 ' End If
    End With
 payGlBillBuy1 = Line1 + 1
ErrTrap:
End Function
 
 Function payGl(LngDevID As Long, notes_id As Double) As Double

If DCboCashType.ListIndex <> 6 Then Exit Function
Dim rsBranch As New ADODB.Recordset
Dim total_value As Double
Dim cProgress As ClsProgress
 Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double
     My_SQL = "SELECT  (branch_id) From TblBranchesData"
   
    rsBranch.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsBranch.RecordCount > 0 Then
        rsBranch.MoveFirst
    End If
    
    
    cProgress.StartProgress

    DoEvents
    total_value = XPTxtVal.text
    
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    'total_value = 0
       Dim Msgdes As String
    Dim BranchID As Integer
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
Dim i As Integer
If SystemOptions.UserInterface = ArabicInterface Then
Msgdes = "»‰«¡ ⁄·Ï „œ›Ê⁄«  —ﬁ„ " & TxtNoteSerial1.text & " "
Msgdes = Msgdes & "”œ«œ —Ê« » ‘Â—"
Msgdes = Msgdes & " " & CmbMonth1.text
Msgdes = Msgdes & " ”‰… " & " " & CboYear1.text
Else
Msgdes = "Payments No. " & TxtNoteSerial1.text
Msgdes = Msgdes & "Payment of salaries of month"
Msgdes = Msgdes & " " & CmbMonth1.text
Msgdes = Msgdes & " year " & " " & CboYear1.text
End If
Dim TotalVacValue As Double
Dim TotalVacValueTotal As Double
OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
    With FrmEmpSalary6.GRID1
Msg = XPMTxtRemarks.text & CHR(13) & txt_general_des & " " & Msgdes


        Dim mLineID As Long
        For i = .FixedRows To .rows - 2
 
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                BranchID = val(.TextMatrix(i, .ColIndex("BranchId")))
                total_value = total_value + Round(.TextMatrix(i, .ColIndex("EmpTotalNet")), 2)
                TotalVacValueTotal = TotalVacValueTotal + Round(val(.TextMatrix(i, .ColIndex("TotalVacValue"))), 2)
                depit_side = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_id"))), "Account_Code1")
                CURRENT_LINE = setfoxy_Line
                TotalVacValue = val(val(.TextMatrix(i, .ColIndex("TotalVacValue"))))
                If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) > 0 Then
                If depit_side = "" Then
                MsgBox ""
                End If
                mLineID = mLineID + 1
                    If ModAccounts.AddNewDev(LngDevID, mLineID, depit_side, Round(.TextMatrix(i, .ColIndex("EmpTotalNet")), 2) - TotalVacValue, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_id"))), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                End If
                If TotalVacValue <> 0 Then
                    depit_side = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_id"))), "Account_code2")
                    mLineID = mLineID + 1
                    If ModAccounts.AddNewDev(LngDevID, mLineID, depit_side, TotalVacValue, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_id"))), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If
                End If
              
                If .TextMatrix(i, .ColIndex("cost_center_id")) <> "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        save_cost_center1 .TextMatrix(i, .ColIndex("cost_center_id")), "”‰œ ﬁÌœ ”œ«œ —« »", XPDtbTrans.value, .TextMatrix(i, .ColIndex("EmpTotalNet")), foxy_ked_NO, depit_side, .TextMatrix(i, .ColIndex("Emp_Name")), CURRENT_LINE
                    Else
                        save_cost_center1 .TextMatrix(i, .ColIndex("cost_center_id")), "Payment Salary JL", XPDtbTrans.value, .TextMatrix(i, .ColIndex("EmpTotalNet")), foxy_ked_NO, depit_side, .TextMatrix(i, .ColIndex("Emp_Name")), CURRENT_LINE
                    End If
                End If
            End If

        Next i

    End With
               
    If total_value > 0 Then
                        
        If getNoOfBranches = 1 Then
                                
     '       If ModAccounts.AddNewDev(LngDevID, i + 1, credit_side, total_value, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, 200, , , , , , , , setfoxy_Line, , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
      
     '           GoTo ErrTrap
     '       End If
                                
        Else '›Ì Õ«·…  ⁄œ «·«›—Ê⁄
            Dim Branch As Integer
            Dim CValue  As Double
Dim DeptSide1 As String
Dim credit_side1 As String

            If rsBranch.RecordCount > 0 Then
                rsBranch.MoveFirst
            End If

            i = i + 1

                                    For Branch = 1 To rsBranch.RecordCount
                                                                                                 
                                        BranchID = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                                                    
                                                    
                                        CValue = GetComponentValuePerBranch2(BranchID, "EmpTotalNet")
                                                   If BranchID = val(Me.dcBranch.BoundText) Then CValue = 0
                                                   
                                                 DeptSide1 = getBranchCurrentAccount(BranchID)
                                                 credit_side1 = getBranchCurrentAccount(dcBranch.BoundText)
                                                    If CValue > 0 Then
                                                                                        
                                                                    If CValue > 0 And DeptSide1 <> "" And credit_side1 <> "" Then
                                                                                    If ModAccounts.AddNewDev(LngDevID, i, DeptSide1, CValue, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                         i = i + 1
                                                                If ModAccounts.AddNewDev(LngDevID, i, credit_side1, CValue, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                                                                        GoTo ErrTrap
                                                                                    End If
                                                                                    
                                                                        i = i + 1
                                                                    End If
                                                                                                                
                                                    End If
                        
                                        rsBranch.MoveNext
                                    Next Branch

        End If
                
    End If
 payGl = i + 1
    With FrmEmpSalary6.GRID1

        For i = .FixedRows To .rows - 2
         
            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) + val(.TextMatrix(i, .ColIndex("OldValue"))) = val(.TextMatrix(i, .ColIndex("NetValue"))) Then
                If Change_filed_value(val(.TextMatrix(i, .ColIndex("id"))), "id", "Payed", "emp_salary", 1) Then
                End If
            End If
            End If

        Next i

    End With

    Dim X As Integer
   
     

    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
ErrTrap:
        
       '     LogTextA = "    ‘«‘…  „”Ì— «·—Ê« »   „ «‰‘«¡ «·ﬁÌœ ··—Ê« » Ê«·„”Ì— " & Chr(13) & " «·‘Â—     " & CmbMonth.text & Chr(13) & "  «·”‰…   " & CboYear.text & Chr(13) & " «· «—ÌŒ " & DTP_Date.value
                     
   ' LogTextE = ""
   '    AddToLogFile CInt(user_id), 555, Date, Time, LogTextA, LogTextE, Me.name, "N", "", , val(TxtNoteSerial), ""
 
 
End Function
 

Function payGlVAT(LngDevID As Long, notes_id As Double) As Double

If DCboCashType.ListIndex <> 12 Then Exit Function
Dim total_value As Double
Dim cProgress As ClsProgress
 Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim Line1 As Double
    cProgress.StartProgress
    DoEvents
    total_value = val(XPTxtVal.text)
    If total_value > 0 Then
  LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    'total_value = 0
    
    Dim BranchID As Integer
    Dim CURRENT_LINE As Double
    Dim depit_side As String
   
   Dim i As Integer
   OtherInformation.NextAccount_Code = DcboCreditSide.BoundText
   BranchID = my_branch
   depit_side = get_account_code_branch(145, my_branch)
         i = i + 1
                CURRENT_LINE = setfoxy_Line
                    If ModAccounts.AddNewDev(LngDevID, i + 1, depit_side, total_value, 0, "Õ”«» ÂÌ∆… «·“ﬂ«…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                        GoTo ErrTrap
                    End If

 payGlVAT = i + 1
End If

    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
ErrTrap:
End Function
Function checkpayrol() As Boolean


checkpayrol = True
Dim i As Integer

   If DCboCashType.ListIndex = 6 Then




                        With FrmEmpSalary6.GRID1

                                       For i = .FixedRows To .rows - 2

                                                   If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then

                                                         GoTo SelectEmp
                                                      End If

                                 Next i

                      End With

    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·„ Ì „  ÕœÌœ «Ì „ÊŸ› ··”œ«œ ·… :"
    Else
        MsgBox " there is No Employee Selected"
    End If
checkpayrol = False
    Exit Function

SelectEmp:


             With FrmEmpSalary6.GRID1

         For i = .FixedRows To .rows - 2

             If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then

                 If get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_id"))), "Account_Code1") = "" Then
                     If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Â‰«ﬂ Œÿ√ ›Ì Õ”«»  «·«ÃÊ— «·„” Õﬁ… ···„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code"))
                    Else
                        MsgBox " Error In Employee Salary Allocation Account For Employee : " & .TextMatrix(i, .ColIndex("Emp_code"))
                     End If
                    Exit Function
                 End If

             End If

          Next i

     End With


  End If

End Function

Private Sub Cmd_Click(Index As Integer)
    Dim cNoteReport As ClsNotesReports
    Dim Msg As String
    'On Error GoTo ErrTrap
 

 
    
    'Option3.value = True
XPTxtVal.text = Format(XPTxtVal.text, "###.00")

    Select Case Index
    Case 15
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments Me.TxtNoteSerial1, "0712201401"
Case 14
   SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
        Case 0


            If SystemOptions.SysRegisterState = DemoRun Then
                If Not rs Is Nothing Then
                    If Not (rs.BOF Or rs.EOF) Then
                        If rs.RecordCount >= 25 Then
                            Msg = "›Ï «·‰”Œ… «· Ã—Ì»Ì… ·«Ì„ﬂ‰  ”ÃÌ· «ﬂÀ— „‰ 25 ⁄„·Ì… ﬁ»÷ «Ê œ›⁄"
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            Exit Sub
                        End If
                    End If
                End If
            End If

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
          
            LblTotalV.Caption = 0
            TxtModFlg.text = "N"
                    TxtChequeNumber.text = "0"
            GRID1.Clear flexClearScrollable, flexClearEverything
GRID1.rows = 1
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
VSFlexGrid1.rows = 1
            VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
VSFlexGrid1.rows = 2
IncludVAT.value = vbChecked
            ' XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=5"))
            Me.DCboUserName.BoundText = user_id
            '      XPDtbTrans.SetFocus
            Text1.text = setfoxy
            Me.dcBranch.BoundText = Current_branch
            Option1.value = False
            Option2.value = False
            Option3.value = False
            Txt_DateHigri.value = ToHijriDate(Date)
XPDtbTrans.SetFocus
XPTxtID1.text = ""
CmbMonth1.ListIndex = 0
CboYear1.text = year(Date)
Me.DcbCurrency.BoundText = MainCurrency()
If CheckAnyVAT(XPDtbTrans.value) = False Then
IncludVAT.value = vbUnchecked
IncludVAT.Enabled = False
Else
IncludVAT.Enabled = True
End If
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
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–… «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
                 Else
                 Msg = " Can't Update "
                    Msg = Msg & CHR(13) & "Cheque Already Payed "
                 End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
   If val(DCboCashType.ListIndex) = 9 Or val(DCboCashType.ListIndex) = 11 Or val(DCboCashType.ListIndex) = 7 Or val(DCboCashType.ListIndex) = 6 Then
   If SystemOptions.UserInterface = ArabicInterface Then
           Msg = "„‰ «Ã·  ⁄œÌ·  «·»Ì«‰«  ”Ì „ Õ–› »Ì«‰«  «·œ›⁄«   " & CHR(13)
           Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    Else
   Msg = " In order to modify the data will be deleted payments data & Chr(13)"
           Msg = Msg + "Confirm Deleted"
     End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
        If val(DCboCashType.ListIndex) = 9 Then
        Command7_Click
        End If
         If val(DCboCashType.ListIndex) = 7 Then
       Command3_Click
        End If
         If val(DCboCashType.ListIndex) = 6 Then
        DeleteSalaryPayment
        End If
          If val(DCboCashType.ListIndex) = 11 Then
        Command11_Click
        End If
        Else
        Exit Sub
   End If
   
   End If
        If CheAdvanced(val(Me.TxtAdvance.text)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–Â «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… —œ ”·›…   "
                    Else
                    Msg = " Can Not Edit this Process"
                    Msg = Msg & CHR(13) & " There is the Process of Advance Payed "
                    
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
     If CheAssetPayd(val(Me.XPTxtID)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–Â «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ≈÷«›… ··«’Ê·   "
                    Else
                    Msg = " Can Not Edit this Process"
                    Msg = Msg & CHR(13) & " There is the Process of adding Assest "
                    
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
  If val(DCboCashType.ListIndex) = 4 Then
   If CheckAdvanecPayed() = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰  ⁄œÌ· Â–Â «·Õ—ﬂ… „— »ÿ… » ⁄œÌ· «·”·›"
   Else
   MsgBox "Can not edit .This process is associated with modification of advances"
   End If
   Exit Sub
   End If
   End If
            TxtModFlg.text = "E"
         '   Me.DCboUserName.BoundText = user_id
            CuurentLogdata
            If CheckAnyVAT(XPDtbTrans.value) = False Then
IncludVAT.value = vbUnchecked
IncludVAT.Enabled = False
Else
IncludVAT.Enabled = True
End If
If val(DCboCashType.ListIndex) = 1 Then
'Command5_Click
Reline2
Reline
RelineBuy
End If
If val(DCboCashType.ListIndex) = 3 Then
'Command5_Click
RelineProject
End If

Dim i As Integer
        Case 2


           If SystemOptions.MonyeIssueVchrNoMust = True And TxtOrder.text = "" And TxtAdvance.text = "" And TxtDue.text = "" Then

                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ   «Œ Ì«— —ﬁ„  «·’—› "
                        Else
                        MsgBox "Please select   Issue Vchr"
                       End If
              Exit Sub
           End If
              
             If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
            ''////////
   If Option5.value = True Then
        If val(CmbMonth.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ  ÕœÌœ «·‘Â—"
        Else
        MsgBox "Please Select Month"
        End If
        CmbMonth.SetFocus
        Exit Sub
        End If
        If val(CboYear.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ  ÕœÌœ «·”‰…"
        Else
        MsgBox "Please Select Year"
        End If
        CboYear.SetFocus
        Exit Sub
        End If
            If ChekPayedSalary(val(CboYear.text), val(CmbMonth.ListIndex) + 1, val(Me.DcbEmpBranch.BoundText)) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ Õ–› ﬁÌœ «·—Ê« »  ··‘Â— «·„Õœœ «Ê·«"
            Else
            MsgBox "Delete Salary Allocation JL"
            End If
            Exit Sub
            End If
        End If
            ''///////
           If val(XPTxtVal.text) = 0 Then
           If SystemOptions.UserInterface = ArabicInterface Then
           MsgBox "Ì—ÃÏ ≈œŒ«· ﬁÌ„… «·„œ›Ê⁄« "
           Else
           MsgBox "Please Enter Value"
           End If
          ' XPTxtVal.SetFocus
           Exit Sub
           End If
           If DCboCashType.ListIndex = 10 Then
           If val(XPTxtVal.text) <= 0 Then
           If SystemOptions.UserInterface = ArabicInterface Then
           MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬁÌ„… «ﬁ· «Ê  ”«ÊÌ «·’›—"
           Else
           MsgBox "Can not be a value less than or equal to zero"
           End If
           Exit Sub
           End If
           End If
        If val(DCboCashType.ListIndex) = 9 Or val(DCboCashType.ListIndex) = 6 Or val(DCboCashType.ListIndex) = 7 Then
        If val(XPTxtVal.text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ ≈Œ Ì«— «·œ›⁄« "
        Else
        MsgBox "Please Select Payments"
        End If
        Exit Sub
        End If
        End If
        

If checkpayrol = False Then Exit Sub

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
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

            '       If Me.TxtModFlg.text = "N" Then
             
            '       End If
            ' TxtNoteSerial.text = Notes_coding(Val(my_branch), XPDtbTrans.value)
    
            If Fra(2).Visible = True Then

                With FG
 If .rows > 1 Then
                    Me.LblTotalV.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("PartValue"), .rows - 1, .ColIndex("PartValue"))
 End If
                End With
    
                If Round(LblTotalV.Caption, 1) <> Round(val(XPTxtVal.text), 1) Then
    
                    If SystemOptions.UserInterface = EnglishInterface Then
                        Msg = "  Instalment Total not correct"
                    Else
                        Msg = "ﬁÌ„… «·œ›⁄«  ·«  ”«ÊÌ «·ﬁÌ„… «·«Ã„«·Ì… ··”‰œ    "
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                   
                    Screen.MousePointer = vbDefault
                    Exit Sub
    
                End If
    
            End If

If val(txtTransferExpenses.text) > 0 And IncludVAT.value = vbChecked Then
If GetValueAddedAccount(XPDtbTrans.value, , , 1, 23) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì…"
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If

If val(Me.TxtPrePayd(17).text) > 0 And (DCboCashType.ListIndex = 7 Or DCboCashType.ListIndex = 1) Then
If GetValueAddedAccount(XPDtbTrans.value, , , 1, 23) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì…"
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If
 If Me.Option1.value = True Then
If val(DCboCashType.ListIndex) = 1 Then
If val(XPTxtVal.text) <= 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· «·ﬁÌ„…"
Else
MsgBox "Please Enter Value "
End If
XPTxtVal.SetFocus
Exit Sub
Else
If Option8(1).value = True Then
RetriveBillVendor val(DBCboClientName.BoundText)
If AutoCalculate2 = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«ÌÊÃœ ›Ê« Ì— «Ê «‰ «·ﬁÌ„… «·„œŒ·… «ﬂ»— „‰ «·„” Õﬁ "
Else
MsgBox "Not Found Bills"
End If
Exit Sub
End If
Else
RetriveBillBuy val(DBCboClientName.BoundText)
If AutoCalculate = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«ÌÊÃœ ›Ê« Ì— «Ê «‰ «·ﬁÌ„… «·„œŒ·… «ﬂ»— „‰ «·„” Õﬁ "
Else
MsgBox "Not Found Bills"
End If
Exit Sub
End If
End If
End If
End If
End If
SaveData
           
           Case 3
'     If val(XPTxtVal.text) = 0 Then
'     If SystemOptions.UserInterface = ArabicInterface Then
'     MsgBox "·«Ì„ﬂ‰ «· —«Ã⁄ Ê«·ﬁÌ„… »’›—"
'     Else
'     MsgBox "You can not undo the value zero "
'     End If
'     ' XPTxtVal.SetFocus
'     Exit Sub
'     End If
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
      
        If CheAdvanced(val(Me.TxtAdvance.text)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–Â «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… —œ ”·›…   "
                    Else
                    Msg = " Can Not Delete this Process"
                    Msg = Msg & CHR(13) & " There is the Process of Advance Payed "
                    
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
        If val(DCboCashType.ListIndex) = 7 Then
 If ChekExpensTotal(PayDes.text) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «·Õ–›  „ «·«ÿ›«¡"
                Else
                MsgBox "Can not Delete"
                End If

Exit Sub
Else


End If

End If

   If val(DCboCashType.ListIndex) = 4 Then
   If CheckAdvanecPayed() = True Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «·Õ–›. Â–Â «·Õ—ﬂ… „— »ÿ… » ⁄œÌ· «·”·›"
   Else
   MsgBox "Can not delete .This process is associated with modification of advances"
   End If
   Exit Sub
   End If
   End If
   
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
            

           Load FrmNotesSearch
            FrmNotesSearch.SearchType = 5
            FrmNotesSearch.show vbModal

        Case 6
        If Me.TxtModFlg.text = "E" Then

        FrmEmpSalary6.VSFlexGrid1.rows = 1

        SaveData
        End If
            Unload Me

        Case 7
                    If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
            If val(DCboCashType.ListIndex) = 9 Or val(DCboCashType.ListIndex) = 6 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Â·  —Ìœ ÿ»«⁄…  Õ·Ì·Ì"
            Else
                Msg = "Do you want to print the analytical"
            End If
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If val(DCboCashType.ListIndex) = 9 Then
            print_reportAnalSuppler
            End If
            If val(DCboCashType.ListIndex) = 6 Then
            print_reportSalary
            End If
            Else
            print_report Me.TxtNoteSerial.text, Me.TxtCustCode.text, val(TxtNoteSerial1.text)
            End If
            Else
                print_report Me.TxtNoteSerial.text, Me.TxtCustCode.text, val(TxtNoteSerial1.text)
             End If
                '     Set cNoteReport = New ClsNotesReports
                '     cNoteReport.PrintReceipt Val(Me.XPTxtID.text), WindowTarget
                '     Set cNoteReport = Nothing
            End If
''''''''''''''''''''''
         '   If DoPremis(Do_Print, Me.name, True) = False Then
         '       Exit Sub
         '   End If
'
'            If val(Me.XPTxtID.text) <> 0 Then
'                print_report Me.TxtNoteSerial.text, Me.TxtCustCode.text
'
'                '     Set cNoteReport = New ClsNotesReports
'                '     cNoteReport.PrintReceipt Val(Me.XPTxtID.text), WindowTarget
'                '     Set cNoteReport = Nothing
'            End If

        Case 9
   
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc Me.TxtNoteSerial.text, , 200, val(XPTxtID.text), txtTotalWithVat
   
        Case 10
   
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_Cheque TxtChequeNumber.text, get_Cheque_report_no(val(DcboBankName.BoundText)), TxtNoteSerial.text
     
        Case 11
            CalCulateParts
            
        Case 12
        called = True
            TxtModFlg.text = "N"
            Me.XPTxtID.text = ""
 XPTxtID1.text = ""
            Me.DCboUserName.BoundText = user_id
              'Me.DcBranch.BoundText = Current_branch
     TxtNoteSerial.text = ""
     TxtNoteSerial1.text = ""
     Text1.text = setfoxy
    called = False
    
    
       Case 13
If TxtReportName.text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— ‰„Ê–Ã  ﬁ—Ì— «·»‰ﬂ „‰ ‘«‘… «·»‰Êﬂ"
Else
MsgBox "Please Select Report From Banks Screen"
End If
Exit Sub
End If
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
              Me.print_reportDeposits
    
            End If
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_reportSalary(Optional NoteSerial As String)
    Dim My_SQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

My_SQL = " SELECT      dbo.TblEmployee.Emp_Name AS Emp_NameH, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, "
My_SQL = My_SQL + "                       dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee1,"
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Namee, dbo.emp_salary.*,"

My_SQL = My_SQL + "                      OldPayment = (Select SUM(PaymentValue) AS SumValue From dbo.TblSalaryNotesPayment Where (EmpID =TblEmployee.Emp_ID) And (YearID = " & val(Me.CboYear1.text) & ") And (MonthID = " & CmbMonth1.ListIndex & " ) and TransID<>" & val(XPTxtID.text) & " )"

My_SQL = My_SQL + " FROM         dbo.emp_salary INNER JOIN"
My_SQL = My_SQL + "                      dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID INNER JOIN"
My_SQL = My_SQL + "                      dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id"
My_SQL = My_SQL + "   WHERE     ( 1=1) "
'My_SQL = My_SQL + "   AND  (payed =1) and     "
My_SQL = My_SQL + "   and (sgn = '" & Me.CboYear1.text & CmbMonth1.ListIndex + 1 & "')   AND (dbo.emp_salary.emp_id in(" & empDes & ")) "
 My_SQL = My_SQL + " order by   ( emp_salary.Emp_code) "

 
 
 

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPaymentSalary.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPaymentSalary.rpt"
        End If
        
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
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
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng

        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue CmbMonth1.text
   ' xReport.ParameterFields(4).AddCurrentValue val(TxtOrderSuppler.text)
    xReport.ParameterFields(5).AddCurrentValue DCPreFix.text & TxtNoteSerial1.text
      xReport.ParameterFields(6).AddCurrentValue txt_general_des.text
       xReport.ParameterFields(7).AddCurrentValue XPDtbTrans.value
       xReport.ParameterFields(8).AddCurrentValue CboYear1.text
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , My_SQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function
Sub SaveSalaryPyment()
Dim i As Integer
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
If Me.TxtModFlg.text = "E" Then

End If
sql = "select * from TblSalaryNotesPayment   where 1=-1"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 With FrmEmpSalary6.GRID1
   For i = 1 To .rows - 1
   If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
      If val(.TextMatrix(i, .ColIndex("Emp_ID"))) <> 0 Then
      Rs3.AddNew
      Rs3("TransID").value = val(XPTxtID.text)
      Rs3("EmpID").value = val(.TextMatrix(i, .ColIndex("Emp_ID")))
      Rs3("PaymentValue").value = val(.TextMatrix(i, .ColIndex("EmpTotalNet")))
      Rs3("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
      Rs3("OldValue").value = val(.TextMatrix(i, .ColIndex("OldValue")))
      Rs3("RemainValue").value = val(.TextMatrix(i, .ColIndex("RemainValue")))
      Rs3("YearID").value = val(CboYear1.text)
      Rs3("MonthID").value = val(CmbMonth1.ListIndex)
      Rs3.update
      End If
    End If
   Next i
End With
End Sub
Function print_reportAnalSuppler(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  MySQL = " SELECT     dbo.TblAttributionContract.IDMC, dbo.TblAttributionContract.ProcessNo, dbo.TblAttributionContract.Name, dbo.TblAttributionContract.Dif, "
MySQL = MySQL & "                      dbo.TblAttributionContract.Depend, dbo.TblAttributionContract.SchoolYear, dbo.TblAttributionContract.FromDate, dbo.TblAttributionContract.FromDateH,"
MySQL = MySQL & "                      dbo.TblAttributionContract.ToDate, dbo.TblAttributionContract.ToDateH, dbo.TblAttributionContract.VendorID, dbo.TblCustemers.CusName,"
MySQL = MySQL & "                      dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankCode, dbo.TblCustemers.BankAddress,"
MySQL = MySQL & "                      dbo.TblCustemers.IBAN, dbo.TblCustemers.BankName, dbo.TblCustemers.BankAccount, dbo.TblCustemers.RecordNo, dbo.TblCustemers.CustGID,"
MySQL = MySQL & "                      dbo.TblAttributionContract.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCustemers.Account_Code,"
MySQL = MySQL & "                      dbo.TblCustemers.CusID, dbo.TblAttributionContract.IDAC, dbo.TblAttributionInstallmentDivided.TotalValue, dbo.TblAttributionInstallmentDivided.ID,"
MySQL = MySQL & "                      dbo.TblAttributionInstallmentDivided.BoardNO , dbo.TblAttributionInstallmentDivided.PayMentPayed"
MySQL = MySQL & " FROM         dbo.TblAttributionInstallmentDivided RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAttributionContract ON dbo.TblAttributionInstallmentDivided.IDAC = dbo.TblAttributionContract.IDAC LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblAttributionContract.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID"
MySQL = MySQL & "   WHERE     (dbo.TblAttributionInstallmentDivided.ID IN ( " & TxtNoSupplerDes.text & "))"
MySQL = MySQL & "  AND  (dbo.TblAttributionInstallmentDivided.PayMentPayed =1)"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPaymentSublerAnal.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPaymentSublerAnal.rpt"
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
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng

        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue dcMontth.text
    xReport.ParameterFields(4).AddCurrentValue val(TxtOrderSuppler.text)
    xReport.ParameterFields(5).AddCurrentValue DCPreFix.text & TxtNoteSerial1.text
      xReport.ParameterFields(6).AddCurrentValue txt_general_des.text
       xReport.ParameterFields(7).AddCurrentValue XPDtbTrans.value
       xReport.ParameterFields(8).AddCurrentValue dcDur.text
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
Private Function CheckDate() As Boolean
    Dim StrTemp As String
    Dim Msg  As String
'
'    If year(Date) > val(Me.CboYear.Text) Then ' ⁄«„ „÷Ï
'        Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ ÕÌÀ «‰Â ﬁ»·  «—ÌŒ «·ÌÊ„ ÊÂ–« ·« ÌÃÊ“...!!!"
''        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
''        CheckDate = False
''        Exit Function
'    If year(Date) = val(Me.CboYear.Text) Then '‰›” «·⁄«„
'
'        If Month(Date) > val(Me.CmbMonth.ListIndex + 1) Then
'            'Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ...!!!"
'            'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            'CheckDate = False
'            'Exit Function
'        End If
'    End If

    CheckDate = True
End Function

Private Function CheckPartCal() As Boolean
    Dim Msg As String

    CheckPartCal = False

    If val(XPTxtVal.text) = 0 Then
        Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·”·›…...!!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 XPTxtVal.SetFocus
        Exit Function
    End If

    If val(TxtPaymentCounts.text) = 0 Then
        Msg = "ÌÃ» «œŒ«· ⁄œœ „—«   ”œÌœ «·œ›⁄…...!!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtPaymentCounts.SetFocus
        Exit Function
    End If

    If CmbMonth.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ «Ê· ‘Â— · ”œÌœ «·œ›⁄…..!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CmbMonth.SetFocus
        Sendkeys "{F4}"
        Exit Function
    End If

    If CboYear.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ «Ê· ”‰… · ”œÌœ «·œ›⁄… ..!! "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboYear.SetFocus
        Sendkeys "{F4}"
        Exit Function
    End If

    CheckPartCal = True
End Function

Private Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2010 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex
    YearMonth1
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

Private Sub CalCulateParts(Optional ByVal FormRet As Boolean = False)
    Dim i As Integer
    Dim IntPartCounts As Integer
    Dim SngPartValue As Double
    Dim m_FirstDate As Date
    If Not FormRet Then
        If CheckPartCal = False Then
            Exit Sub
        End If
    
        If CheckDate = False Then
            Exit Sub
        End If
    End If
    
    SngPartValue = val(Me.XPTxtVal.text) / val(Me.TxtPaymentCounts.text)
    IntPartCounts = val(Me.TxtPaymentCounts.text)
    m_FirstDate = CDate(val(Me.CboYear.text) & "-" & Me.CmbMonth.ListIndex + 1 & "-01")

    With Me.FG
        .Clear flexClearScrollable, flexClearEverything
        .rows = .FixedRows + IntPartCounts
        .RowHeightMin = 300

        For i = 1 To IntPartCounts
            .TextMatrix(i, .ColIndex("PartNO")) = i
            .TextMatrix(i, .ColIndex("PartValue")) = Round(SngPartValue, 2)
            .TextMatrix(i, .ColIndex("PartDate")) = DisplayDate(DateAdd("m", i - 1, m_FirstDate))
        Next i
    
    
                
 
                    Me.LblTotalV.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("PartValue"), .rows - 1, .ColIndex("PartValue"))
         
                
                
    End With

End Sub
Public Function print_report(Optional NoteSerial As String, Optional Custcode As String, Optional NoteSerial1 As Double = 0)
    On Error Resume Next
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If NoteSerial1 = 0 Then
    MySQL = "Select *, '" & TxtCustCode & "' as CustCode, EXPENSES_ORDER2.TransferExpenses, '" & DcboDebitSide.text & "' as DebitSide,'" & DcboCreditSide.text & "' as CreditSide ,TblCustemers.* From EXPENSES_ORDER2  "
    MySQL = MySQL & " Left outer join TblCustemers On  TblCustemers.CusId = EXPENSES_ORDER2.CusId "
    MySQL = MySQL & "where NoteSerial='" & NoteSerial & "'"
    
 Else
 
     MySQL = "Select    '" & Trim(lbl(18).Caption) & "' as VallueChr, '" & Trim(TxtCustCode) & "' as CustCode,EXPENSES_ORDER2.TransferExpenses, Note_Value2 as Note_Value,EXPENSES_ORDER2.*,TblCustemers.*,'" & DcboDebitSide.text & "' as DebitSide,'" & DcboCreditSide.text & "' as CreditSide "
MySQL = MySQL & "       ,Note_Value2 AS Note_Value"
   MySQL = MySQL & " ,Note_Value2 AS Note_ValueVV"
   MySQL = MySQL & " ,Note_Value AS Note_ValueVV1"
     MySQL = MySQL & " From EXPENSES_ORDER2"
     
     MySQL = MySQL & " Left outer join TblCustemers On  TblCustemers.CusId = EXPENSES_ORDER2.CusId "
     MySQL = MySQL & " where  NoteSerial1=" & NoteSerial1 & " and notetype=5"
     MySQL = MySQL & " and EXPENSES_ORDER2.NoteID = " & XPTxtID.text
 End If
 Dim s As String
 If Option5.value = True Then
        s = " SELECT dbo.BanksData.BankName,'" & DcboDebitSide.text & "' as DebitSide,'" & DcboCreditSide.text & "' as CreditSide,"
        s = s & "       Notes.AdvanceID,"
        s = s & "       dbo.Notes.NoteID,"
        s = s & "                dbo.Notes.NoteDate,"
        s = s & "                dbo.Notes.NoteType,"
        s = s & "                dbo.Notes.NoteSerial,"
        s = s & "                dbo.Notes.Note_Value,"
        s = s & "                dbo.Notes.BankID,"
        s = s & "                dbo.Notes.ChqueNum,"
        s = s & "                dbo.Notes.DueDate,"
        s = s & "                dbo.Notes.NoteHijriDate,"
        s = s & "                dbo.Notes.Transaction_ID,"
        s = s & "                dbo.Notes.MaintananceID,"
        s = s & "                dbo.Notes.Member_ID,"
        s = s & "                dbo.Notes.UserID,"
        s = s & "                dbo.Notes.Remark,"
        s = s & "                dbo.Notes.CashingType,"
        s = s & "                dbo.Notes.ExpensesID,"
        s = s & "                dbo.Notes.BoxID,"
        s = s & "                dbo.Notes.CusID,"
        s = s & "                dbo.Notes.RetrunNoteID,"
        s = s & "                dbo.Notes.RevenuesID,"
        s = s & "                dbo.Notes.NotePosted,"
        s = s & "                dbo.Notes.NoteCashingType,"
        s = s & "                dbo.Notes.PostedBy,"
        s = s & "                dbo.Notes.PostDate,"
        s = s & "                dbo.Notes.NumOrderInpot,"
        s = s & "                dbo.Notes.Buy,"
        s = s & "                dbo.Notes.ked_type,"
        s = s & "                dbo.Notes.numbering_type,"
        s = s & "                dbo.Notes.sanad_year,"
        s = s & "                dbo.Notes.sanad_month,"
        s = s & "                dbo.Notes.type,"
        s = s & "                dbo.Notes.branch_no,"
        s = s & "                dbo.Notes.user_name,"
        s = s & "                dbo.Notes.DEPARTEMENT,"
        s = s & "                dbo.Notes.sanad_type,"
        s = s & "                dbo.Notes.sanad_source,"
        s = s & "                dbo.Notes.Double_Entry_Vouchers_ID,"
        s = s & "                dbo.Notes.DAWRY,"
        s = s & "                dbo.Notes.KALEB,"
        s = s & "                dbo.Notes.projectAccountCode,"
        s = s & "                dbo.Notes.foxy_no,"
        s = s & "                dbo.Notes.person,"
        s = s & "                dbo.Notes.project_Expensen_account,"
        s = s & "                dbo.Notes.salary,"
        s = s & "                dbo.Notes.displayed,"
        s = s & "                dbo.Notes.Adv_payment_value,"
        s = s & "                dbo.Notes.note_value_by_characters,"
        s = s & "                dbo.Notes.too,"
        s = s & "                dbo.Notes.notes_all,"
        s = s & "                dbo.Notes.general_cost_center,"
        s = s & "                dbo.Notes.NoteSerial1,"
        s = s & "                dbo.Notes.Cus_or_sub,"
        s = s & "                dbo.Notes.project_id,"
        s = s & "                dbo.Notes.numbering_type1,"
        s = s & "                dbo.Notes.project_depit_or_credit,"
        s = s & "                dbo.Notes.salary_or_advance,"
        s = s & "                dbo.Notes.general_des_notes,"
        s = s & "                dbo.Notes.DeptID,"
        s = s & "                dbo.TblEmpDepartments.DepartmentName,"
        s = s & "                dbo.TblEmpDepartments.DepartmentNamee,"
        s = s & "                dbo.Notes.TxtChequeNumber1,"
        s = s & "                dbo.Notes.DrawingType,"
        s = s & "                  dbo.BanksData.BankNamee,"
        s = s & "                dbo.TblBranchesData.branch_name,"
        s = s & "                dbo.TblBranchesData.branch_namee,"
        s = s & "                dbo.Notes.PreVAT,"
        s = s & "                dbo.Notes.IncludVAT,"
        s = s & "                dbo.Notes.TotalValue,"
        s = s & "                dbo.Notes.VATYou,"
        s = s & "                dbo.Notes.VAT,"
        s = s & "                dbo.Notes.VATVowalNo,"
        s = s & "                dbo.Notes.LockSalary,"
        s = s & "                dbo.Notes.TransferExpenses,"
        s = s & "                dbo.Notes.Note_Value2,"
        s = s & "                dbo.Notes.note_value_by_characters2,"
        s = s & "                dbo.Notes.BeneficiaryBanck,"
        s = s & "                dbo.Notes.BenefIBAN,"
        s = s & "                TblEmpAdvanceDetails.PartNo , TblEmpAdvanceDetails.PartValue, TblEmpAdvanceDetails.PartDate, TblEmpAdvanceDetails.payed"
        s = s & "         From dbo.Notes"
        s = s & "                LEFT OUTER JOIN dbo.TblBranchesData"
        s = s & "                     ON  dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
        s = s & "                LEFT OUTER JOIN dbo.TblEmpDepartments"
        s = s & "                     ON  dbo.Notes.DeptID = dbo.TblEmpDepartments.DeparmentID"
        s = s & "                LEFT OUTER JOIN dbo.BanksData"
        s = s & "                     ON  dbo.Notes.BankID = dbo.BanksData.BankID"
        s = s & "                LEFT OUTER JOIN TblEmpAdvanceDetails"
        s = s & "                     ON  TblEmpAdvanceDetails.AdvanceID = Notes.AdvanceID"
        s = s & "                     Where NoteSerial1 = " & NoteSerial1 & " And NoteType = 5"
        s = s & " and Notes.NoteID = " & XPTxtID.text
        
        MySQL = s
    End If
 
 
    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If
'
'Expenses_order3
If SystemOptions.PaymentDifferent = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
     '    StrFileName = App.path & "\Reports\" & "PaymentVoucher.rpt"
      StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PaymentVoucher.rpt"
          
      Else
        '  StrFileName = App.path & "\Reports\" & "PaymentVoucher_Eng.rpt"
             StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PaymentVoucher.rpt"
      End If
    
  Else
        If CboPayMentType.ListIndex = 0 Then
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PaymentCash.rpt"
    Else
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PaymentCheque.rpt"
    End If
    
 
 End If
 If Option5.value = True Then
    StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PaymentVoucherByLoan.rpt"
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
    If DCboCashType.ListIndex = 3 Then
        
        Dim rsDummy As New ADODB.Recordset
         s = "SELECT p.Project_name,p.Project_nameE FROM projects AS p WHERE p.id = " & val(DBCboClientName.BoundText)

        rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly, adCmdText
        If Not rsDummy.EOF Then
            xReport.ParameterFields(8).AddCurrentValue rsDummy!Project_name & ""
            xReport.ParameterFields(9).AddCurrentValue rsDummy!Project_nameE & ""
        End If
    End If

    
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(5).AddCurrentValue DcboDebitSide.text  'RPTCompany_Name_Arabic
        xReport.ParameterFields(6).AddCurrentValue Custcode
        xReport.ParameterFields(7).AddCurrentValue DcboCreditSide.text  'RPTCompany_Name_Arabic   xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text  'RPTCompany_Name_Arabic
        'xReport.ParameterFields(8).AddCurrentValue DBCboClientName.Text
        'CustCode
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
        xReport.ParameterFields(5).AddCurrentValue DcboDebitSide.text
        xReport.ParameterFields(6).AddCurrentValue Custcode
        xReport.ParameterFields(7).AddCurrentValue DcboCreditSide.text
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Public Function print_reportDeposits(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "SELECT       '" & DcbCurrency.text & "' as CurrencyName ,Notes.note_value_by_characters,  dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value, dbo.Notes.BenefiBanckCode, "
    MySQL = MySQL & "                  dbo.Notes.BenefiBanckAddress, dbo.Notes.BeneficiaryACNo, dbo.Notes.BeneficiaryAddress, dbo.Notes.BeneficiaryBanck, dbo.Notes.NumIqama,"
    MySQL = MySQL & "                  dbo.Notes.Telephone, dbo.Notes.person, dbo.Notes.RemitterName, dbo.Notes.PaymentType, dbo.Notes.BankID, dbo.Notes.ChqueNum, dbo.BanksData.BankName,"
    MySQL = MySQL & "                  dbo.BanksData.Remarks, dbo.BanksData.ReportName, dbo.BanksData.AccountName, dbo.BanksData.Currency, dbo.BanksData.BranchName,"
    MySQL = MySQL & "                  dbo.BanksData.Commision, dbo.BanksData.BankNamee, dbo.BanksData.IBan, dbo.BanksData.Tel, dbo.BanksData.Address, dbo.BanksData.Email,"
    MySQL = MySQL & "                  dbo.BanksData.chkapprov, dbo.BanksData.chkLoan, dbo.BanksData.account_no, dbo.BanksData.Currency_ID, dbo.currency.code, dbo.currency.name,"
    MySQL = MySQL & "                  dbo.currency.nameE, dbo.Notes.KafeltEL, dbo.Notes.KafelName, dbo.Notes.Adress2, dbo.Notes.Street, dbo.Notes.City, dbo.Notes.Governorate, dbo.Notes.Country,"
    MySQL = MySQL & "                  dbo.Notes.BenefGovernorate, dbo.Notes.BenefStreet, dbo.Notes.BenefCity, dbo.Notes.BenefCountry, dbo.Notes.BenefNumIqama, dbo.Notes.DueDate,"
    MySQL = MySQL & "                  dbo.Notes.Remark, dbo.Notes.CashingType, dbo.Notes.NoteCashingType, dbo.Notes.PostDate, dbo.Notes.TransferExpenses, dbo.Notes.ExpensesRemark,"
    MySQL = MySQL & "                  dbo.Notes.RemarkE, dbo.Notes.BenefTelephone, dbo.Notes.BenefPlaceBrith, dbo.Notes.BenefBrithDate, dbo.Notes.BenefDateExpEqama, dbo.Notes.KafelAddress,"
    MySQL = MySQL & "                  dbo.Notes.BenefPlaceIqama , dbo.Notes.BenefIBAN,"
    MySQL = MySQL & "                  '" & DcbCurrency.BoundText & "' as CurrencyID "
    MySQL = MySQL & "   FROM         dbo.currency RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.BanksData ON dbo.currency.id = dbo.BanksData.Currency_ID RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.Notes ON dbo.BanksData.BankID = dbo.Notes.BankID"
    MySQL = MySQL & ""
    MySQL = MySQL & "   Where (dbo.Notes.NoteType = 5) And (dbo.Notes.NoteID =" & val(XPTxtID.text) & ")"

    If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\Deposits\" & TxtReportName.text
    Else
         StrFileName = App.path & "\REPORTS\Deposits\" & TxtReportName.text
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
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
'
       ' xReport.ParameterFields(7).AddCurrentValue DcboCreditSide.text  'RPTCompany_Name_Arabic   xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text  'RPTCompany_Name_Arabic
       ' StrReportTitle = "" '& StrAccountName
    Else
         ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
       ' xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
       ' xReport.ParameterFields(5).AddCurrentValue DcboDebitSide.text
       ' xReport.ParameterFields(6).AddCurrentValue Custcode
       ' StrReportTitle = ""
    End If
xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
    xReport.ParameterFields(3).AddCurrentValue user_name
     xReport.ParameterFields(4).AddCurrentValue WriteNo(XPTxtVal.text, 0)
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
        ' ›« Ê—… „‘ —Ì« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = PurchaseTransaction
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPayMentType.ListIndex = 1
        FrmBuySearch.CboPayMentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… ‘—«¡"
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show
    ElseIf Me.CboTrans.ListIndex = 1 Then
        '›« Ê—… „— Ã⁄ „»Ì⁄« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = ReturnSalling
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPayMentType.ListIndex = 1
        FrmBuySearch.CboPayMentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ „— Ã⁄ «·„»Ì⁄« "
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show vbModal
    End If

End Sub

Private Sub Command1_Click()
Unload FrmEmpSalary6
Load FrmEmpSalary6
FrmEmpSalary6.show
FrmEmpSalary6.Command4.Visible = False
FrmEmpSalary6.Command3.Visible = True
FrmEmpSalary6.Frame3.Visible = True
FrmEmpSalary6.Check22.Visible = False
FrmEmpSalary6.CmbMonth.text = CmbMonth1.text
FrmEmpSalary6.CboYear.text = CboYear1.text
 FrmEmpSalary6.empDes = empDes.text
FrmEmpSalary6.GRID1.Visible = True
FrmEmpSalary6.Check17.Visible = True
FrmEmpSalary6.lbl(12).Visible = True
FrmEmpSalary6.DTPicker1.Visible = True
FrmEmpSalary6.ALLButton3.Visible = True
FrmEmpSalary6.ALLButton6.Visible = False
FrmEmpSalary6.Check18.Visible = False
FrmEmpSalary6.Check19.Visible = True
FrmEmpSalary6.VSFlexGrid1.Visible = False
FrmEmpSalary6.FillGridWithData2
FrmEmpSalary6.ALLButton8.Visible = False
FrmEmpSalary6.VSFlexGrid3.Visible = False
FrmEmpSalary6.Check21.Visible = False
FrmEmpSalary6.Ele(0).Visible = False
FrmEmpSalary6.ALLButton7.Visible = False
FrmEmpSalary6.Check21.Visible = False
FrmEmpSalary6.Check20.Visible = False
End Sub
Public Function DeletePayedSalary(sgn1 As String, empDes1 As String)
If empDes1 = "" Then Exit Function
Dim My_SQL As String
My_SQL = "update emp_salary"
 
My_SQL = My_SQL & "  Set payed = 0"
'My_SQL = My_SQL & "   Where       (sgn = '" & Me.CboYear.text & CmbMonth.ListIndex + 1 & "')   AND (dbo.emp_salary.emp_id in(" & empDes & ")) "
My_SQL = My_SQL & "   Where       (sgn = '" & sgn1 & "')   AND (dbo.emp_salary.emp_id in(" & empDes1 & ")) "
Cn.Execute My_SQL
empDes.text = ""
End Function
Public Function DeletePayedPayment2(OrderSupplerDes1 As String)
Dim My_SQL As String
If OrderSupplerDes1 = "" Then Exit Function
My_SQL = " update TblAttributionInstallmentDivided Set NoteSerial1 =null,noteid=null, PayMentPayed =Null  Where   (ID in (" & OrderSupplerDes1 & "))"
Cn.Execute My_SQL
End Function

Private Function save_cost_center1(cost_center_id As String, _
                                  opr_type As String, _
                                  record_date As Date, _
                                  value As Double, _
                                  kedno As String, _
                                  account_no As String, _
                                  account_name As String, _
                                  line_no As Double)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
    rs("cost_center_id").value = cost_center_id
    rs("cost_center").value = get_EMPLOYEE_COST_CENTER_NAME(cost_center_id, "ACCOUNT_NAME")
    rs("value").value = value
    rs("depit_or_credit").value = "„œÌ‰"
    rs("opr_id").value = kedno
    rs("kedno").value = kedno
        
    rs("opr_type").value = opr_type
    rs("account_name").value = account_name
    rs("account_no").value = account_no
    rs("line_no").value = line_no
    rs("record_date").value = record_date
    rs.update
    rs.Close

End Function

Private Sub Command2_Click()
DeleteSalaryPayment 1
End Sub
Sub DeleteSalaryPayment(Optional Ind As Integer = 0)
XPTxtVal.text = 0
FrmEmpSalary6.ClearSalary = True
DeletePayedSalary Me.CboYear1.text & CmbMonth1.ListIndex + 1, empDes
If Ind = 1 Then
Cn.Execute "Delete from TblSalaryNotesPayment where TransID=" & val(XPTxtID.text) & ""
End If
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·€«¡ «·”œ«œ"
Else
MsgBox "Done"
End If
End Sub
Private Sub Command3_Click()

If ChekExpens(PayDes.text) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ ≈·€«¡ «·”œ«œ"
        Else
                 MsgBox "Can not Cancel "
        End If
        Exit Sub
End If

XPTxtVal.text = 0
FrmEmpSalary6.ClearPayment = True
DeletePayedPayment PayDes.text
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·€«¡ «·”œ«œ"
Else
MsgBox "Done"

End If
End Sub
Public Function DeleteMofrdPayment(PayDes1 As String, empDes1 As String)
Dim My_SQL As String
If PayDes1 = "" Then Exit Function
My_SQL = " update TblApproveCompoYearDet Set PaymentPayed = null  Where   (id in (" & PayDes1 & "))"
Cn.Execute My_SQL
If empDes1 = "" Then Exit Function
My_SQL = " update TblComponentYearDet Set PaymentPayed = null  Where   (id in (" & empDes1 & "))"
Cn.Execute My_SQL
End Function

Public Function DeletePayedPayment(PayDes1 As String)
Dim My_SQL As String
If PayDes1 = "" Then Exit Function
My_SQL = " update TblPripaidExpensesDet Set PaymentPayed = 0  Where   (id in (" & PayDes1 & "))"
                  
Cn.Execute My_SQL
End Function

Public Function DeletePayedPaymeQest()
Dim My_SQL As String
Dim PayDes1 As String
Dim i As Integer
With GRID1
For i = 1 To .rows - 1
PayDes1 = .TextMatrix(i, .ColIndex("StrQest"))
If PayDes1 <> "" Then
My_SQL = " update TblQestFexed Set FlgPaye = Null  Where   (QestID in (" & PayDes1 & "))"
Cn.Execute My_SQL
       My_SQL = "update  notes_all set FlgPaye = Null  Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
    Cn.Execute My_SQL, , adExecuteNoRecords
End If
Next i
End With
End Function
Private Sub Command4_Click()
Unload FrmEmpSalary6
Load FrmEmpSalary6
FrmEmpSalary6.show
FrmEmpSalary6.Check22.Visible = False
FrmEmpSalary6.Frame3.Visible = False
FrmEmpSalary6.Command3.Visible = False
FrmEmpSalary6.Command4.Visible = False
FrmEmpSalary6.ALLButton8.Visible = False
FrmEmpSalary6.VSFlexGrid3.Visible = False
FrmEmpSalary6.Check21.Visible = False
FrmEmpSalary6.PayDes = PayDes.text
FrmEmpSalary6.ALLButton3.Visible = False
FrmEmpSalary6.ALLButton6.Visible = True
FrmEmpSalary6.Grid2.Visible = False
FrmEmpSalary6.GRID1.Visible = False
FrmEmpSalary6.Check17.Visible = False
FrmEmpSalary6.lbl(12).Visible = False
FrmEmpSalary6.DTPicker1.Visible = False
FrmEmpSalary6.VSFlexGrid1.Visible = True
FrmEmpSalary6.Check18.Visible = True
FrmEmpSalary6.Check19.Visible = False
FrmEmpSalary6.ALLButton7.Visible = False
FrmEmpSalary6.Check20.Visible = False
FrmEmpSalary6.VSFlexGrid2.Visible = False
FrmEmpSalary6.FillGrid4
FrmEmpSalary6.Check18.value = vbUnchecked
       Dim i As Integer
        With FrmEmpSalary6.VSFlexGrid1

            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("ch")) = False
            Next i

        End With
        
        
End Sub

Sub RetriveBillProject(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
Set Rs8 = New ADODB.Recordset
sql = " SELECT     TOP 100 PERCENT dbo.project_billl.id, dbo.project_billl.bill_date, dbo.project_billl.project_no, dbo.projects.Project_name, dbo.projects.Fullcode, "
sql = sql & "                       dbo.projects.Project_nameE, dbo.project_billl.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.project_billl.total,"
sql = sql & "                       dbo.project_billl.ManualNO , dbo.project_billl.subContractorId ,dbo.project_billl.NoteSerial1 "
sql = sql & "  FROM         dbo.project_billl LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData ON dbo.project_billl.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                       dbo.projects ON dbo.project_billl.project_no = dbo.projects.id"
sql = sql & " Where (dbo.project_billl.subContractorId = " & CuID & ")and dbo.project_billl.bill_to=1  and dbo.project_billl.project_no=" & ProjectIDD & " And( (dbo.project_billl.totalPayed Is Null)or dbo.project_billl.totalPayed=0) "
sql = sql & "  ORDER BY dbo.project_billl.bill_date"

Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
VSFlexGrid2.Enabled = True
        VSFlexGrid2.Enabled = True
With VSFlexGrid2
.Clear flexClearScrollable, flexClearEverything
.rows = 1
    .rows = .rows + Rs8.RecordCount
.rows = .FixedRows + Rs8.RecordCount
Rs8.MoveFirst
For i = .FixedRows To Rs8.RecordCount
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("Branch_NO").value), 0, Rs8("Branch_NO").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_name").value), 0, Rs8("Project_name").value)
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
Else
.TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(Rs8("Project_nameE").value), 0, Rs8("Project_nameE").value)
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
End If
.TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
.TextMatrix(i, .ColIndex("project_no")) = IIf(IsNull(Rs8("project_no").value), 0, Rs8("project_no").value)

.TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs8("id").value), 0, Rs8("id").value)
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("bill_date").value), Date, Rs8("bill_date").value)
'.TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs8("id").value), "", Rs8("id").value)
.TextMatrix(i, .ColIndex("too")) = IIf(IsNull(Rs8("ManualNO").value), "", Rs8("ManualNO").value)
.TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(Rs8("total").value), 0, Rs8("total").value)
If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteBillProject(val(.TextMatrix(i, .ColIndex("NoteID"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("Note_Value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
Rs8.MoveNext
Next i
End With
End If

End Sub
Sub RetriveBillBuy(Optional CuID As Double = 0, Optional ByVal TradingContractID As Integer = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
Set Rs8 = New ADODB.Recordset

sql = " SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1,"
sql = sql & "                       dbo.Transactions.ManualNO, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.CusID,"
sql = sql & "                       dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.TotalPayed, dbo.Transactions.Transaction_NetValue,"
sql = sql & "                       dbo.Transactions.Currency_rate,dbo.Transactions.Currency_id,currency.name as currencyName,"
sql = sql & "                       dbo.Transactions.DueDate , dbo.Transactions.OldValue"
sql = sql & "  FROM         dbo.Transactions LEFT OUTER JOIN"
sql = sql & "                       dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & "                       Left outer join currency On Transactions.Currency_id = currency.id"
sql = sql & " WHERE     (dbo.Transactions.Transaction_Type = 22 or dbo.Transactions.Transaction_Type = 73) AND (dbo.Transactions.TotalPayed IS NULL OR"
sql = sql & "                       dbo.Transactions.TotalPayed = 0) "
If TradingContractID <> 0 Then
    sql = sql & "                       AND (dbo.Transactions.poTransaction_ID = " & TradingContractID & ")"
Else
    sql = sql & "                       AND (dbo.Transactions.CusID = " & CuID & ")"
End If
sql = sql & "  ORDER BY dbo.Transactions.DueDate"

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
.TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs8("Transaction_ID").value), 0, Rs8("Transaction_ID").value)
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("Transaction_Date").value), Date, Rs8("Transaction_Date").value)
.TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
.TextMatrix(i, .ColIndex("too")) = IIf(IsNull(Rs8("ManualNO").value), "", Rs8("ManualNO").value)
.TextMatrix(i, .ColIndex("Note_ValueE")) = IIf(IsNull(Rs8("Transaction_NetValue").value), IIf(IsNull(Rs8("OldValue").value), 0, Rs8("OldValue").value), Rs8("Transaction_NetValue").value)
.TextMatrix(i, .ColIndex("currencyName")) = IIf(IsNull(Rs8("currencyName").value), "", Rs8("currencyName").value)
.TextMatrix(i, .ColIndex("Currency_rate")) = IIf(IsNull(Rs8("Currency_rate").value), "", Rs8("Currency_rate").value)
.TextMatrix(i, .ColIndex("Note_Value")) = val(.TextMatrix(i, .ColIndex("Note_ValueE"))) * IIf(val(.TextMatrix(i, .ColIndex("Currency_rate"))) <> 0, val(.TextMatrix(i, .ColIndex("Currency_rate"))), 1)


'.TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(Rs8("Transaction_NetValue").value), IIf(IsNull(Rs8("OldValue").value), 0, Rs8("OldValue").value), Rs8("Currency_rate").value)
.TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(Rs8("DueDate").value), Date, Rs8("DueDate").value)
If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteBillBuy(val(.TextMatrix(i, .ColIndex("NoteID"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("Note_Value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
.TextMatrix(i, .ColIndex("RemainingValueE")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) / IIf(val(.TextMatrix(i, .ColIndex("Currency_rate"))) <> 0, val(.TextMatrix(i, .ColIndex("Currency_rate"))), 1)

Rs8.MoveNext
Next i
End With
End If

End Sub
Sub RetriveBillVendor(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
Set Rs8 = New ADODB.Recordset
sql = "SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteSerial1, dbo.notes_all.too, dbo.notes_all.Note_Value, dbo.notes_all.NoteType, dbo.notes_all.CusID, "
sql = sql & "                      dbo.notes_all.totalPayed , dbo.notes_all.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE ,dbo.notes_all.FlgQst ,dbo.notes_all.FATValue"
sql = sql & " FROM         dbo.notes_all LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
sql = sql & " WHERE     (dbo.notes_all.NoteType = 80) AND (dbo.notes_all.CusID = " & CuID & ") AND (dbo.notes_all.TotalPayed IS NULL OR"
sql = sql & "                      dbo.notes_all.TotalPayed = 0)"

sql = sql & " order by  dbo.notes_all.NoteDate"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GRID1.Enabled = True
        Check18.Enabled = True
With GRID1
.Clear flexClearScrollable, flexClearEverything
.rows = 1
    .rows = .rows + Rs8.RecordCount
.rows = .FixedRows + Rs8.RecordCount
Rs8.MoveFirst
For i = .FixedRows To Rs8.RecordCount
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("branch_no").value), 0, Rs8("branch_no").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
Else
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
End If
If Rs8("FlgQst").value = True Then
.TextMatrix(i, .ColIndex("haveqest")) = True
Else
.TextMatrix(i, .ColIndex("haveqest")) = False
End If
.TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs8("NoteID").value), 0, Rs8("NoteID").value)
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("NoteDate").value), Date, Rs8("NoteDate").value)
.TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
.TextMatrix(i, .ColIndex("too")) = IIf(IsNull(Rs8("too").value), "", Rs8("too").value)
.TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(Rs8("FATValue").value), 0, Rs8("FATValue").value) + IIf(IsNull(Rs8("Note_Value").value), 0, Rs8("Note_Value").value)
If val(.TextMatrix(i, .ColIndex("NoteSerial1"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteBillVendor(val(.TextMatrix(i, .ColIndex("NoteSerial1"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("Note_Value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
'.TextMatrix(i, .ColIndex("RemainingValueE")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) * IIf(val(.TextMatrix(i, .ColIndex("Currency_rate"))) <> 0, val(.TextMatrix(i, .ColIndex("Currency_rate"))), 1)
Rs8.MoveNext
Next i
End With
End If

End Sub
Function GeteBillProject(Optional Transaction_ID As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS Smatiobn"
sql = sql & " From dbo.TblBillProjectPayment"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
sql = sql & " GROUP BY Transaction_ID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteBillProject = IIf(IsNull(Rs8("Smatiobn").value), 0, Rs8("Smatiobn").value)
Else
GeteBillProject = 0
End If
End Function
Function GeteBillBuy(Optional Transaction_ID As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS Smatiobn"
sql = sql & " From dbo.TblBillBuyPayment"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
sql = sql & " GROUP BY Transaction_ID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteBillBuy = IIf(IsNull(Rs8("Smatiobn").value), 0, Rs8("Smatiobn").value)
Else
GeteBillBuy = 0
End If
End Function

Function GeteBillVendor(Optional NoteSerial1 As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT     NoteSerial1, SUM(PayedValue) AS Smatiobn"
sql = sql & " From dbo.TblBillVindorPayment"
sql = sql & " Where (NoteSerial1 = " & NoteSerial1 & ")"
sql = sql & " GROUP BY NoteSerial1"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteBillVendor = IIf(IsNull(Rs8("Smatiobn").value), 0, Rs8("Smatiobn").value)
Else
GeteBillVendor = 0
End If
End Function

Private Sub Command6_Click()
Dim i As Integer
Dim StrSQL As String
If Me.TxtModFlg.text = "E" Then
DeleteBill
GRID1.Enabled = True
        Check18.Enabled = True
      StrSQL = "Delete From TblNotesBillVindorPayment Where NoteID1=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillVindorPayment Where NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    DeletePayedPaymeQest
XPTxtVal.text = 0
            GRID1.Clear flexClearScrollable, flexClearEverything
GRID1.rows = 1

FlgBill = True
 FrmEmpSalary6.ClearPayment1 = True
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·€«¡ «·”œ«œ"
Else
MsgBox "Done"
End If
    With Me.GRID1

            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i


        End With
End If
End Sub

Private Sub Command7_Click()
XPTxtVal.text = 0
FrmEmpSalary6.ClearPayment = True
DeletePayedPayment2 TxtNoSupplerDes.text
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·€«¡ «·”œ«œ"
Else
MsgBox "Done"
End If

End Sub
Function checkReq() As Boolean
Dim Rs6 As ADODB.Recordset
Dim sql As String
Set Rs6 = New ADODB.Recordset
checkReq = False
sql = "Select EntryCreated  from TblExchangeRequest where EntryCreated=1 and ID= " & val(TxtOrderSuppler.text) & ""
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
checkReq = True
Else
checkReq = False
End If
End Function
Private Sub Command8_Click()
Dim YraID As Integer
Dim MonthID As Integer
Dim BranchID As Integer
If checkReq() = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „ «‰‘«¡ ﬁÌœ ·Â–« «·ÿ·»"
Else
MsgBox "It has not been set up entry this request"
End If
Exit Sub
Else
Dim AllID As String
If val(TxtOrderSuppler.text) <> 0 Then
Unload FrmEmpSalary6
Load FrmEmpSalary6
FrmEmpSalary6.show

FrmEmpSalary6.Check22.Visible = False
FrmEmpSalary6.Command4.Visible = True
FrmEmpSalary6.Command3.Visible = False
FrmEmpSalary6.ALLButton8.Visible = False
FrmEmpSalary6.VSFlexGrid3.Visible = False
FrmEmpSalary6.Check21.Visible = False
FrmEmpSalary6.Frame3.Visible = False
FrmEmpSalary6.OrderSupplerDes = TxtNoSupplerDes.text
FrmEmpSalary6.ALLButton3.Visible = False
FrmEmpSalary6.ALLButton6.Visible = False
FrmEmpSalary6.Grid2.Visible = False
FrmEmpSalary6.GRID1.Visible = False
FrmEmpSalary6.Check17.Visible = False
FrmEmpSalary6.lbl(12).Visible = False
FrmEmpSalary6.DTPicker1.Visible = False
FrmEmpSalary6.VSFlexGrid1.Visible = False
FrmEmpSalary6.Check18.Visible = False
FrmEmpSalary6.Check19.Visible = False
FrmEmpSalary6.ALLButton7.Visible = True
FrmEmpSalary6.Check20.Visible = True
AllID = GetExchangReq(val(TxtOrderSuppler.text), YraID, MonthID, BranchID)
dcDur.BoundText = YraID
dcMontth.BoundText = MonthID
DcbBrReq.BoundText = BranchID
FrmEmpSalary6.VSFlexGrid2.Visible = True
If AllID <> "" Then
FrmEmpSalary6.FillGrid5 AllID
End If
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— —ﬁ„ ÿ·» ’—› «·„ ⁄ÂœÌ‰"
Else
MsgBox "Please Eneter Number of Reques"
End If
TxtOrderSuppler.SetFocus
Exit Sub
End If
End If
End Sub
Function GetEmployeeBranch(EmpID As Double) As Integer
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "SELECT     Emp_ID, BranchId"
sql = sql & " From dbo.TblEmployee"
sql = sql & " WHERE     (Emp_ID = " & EmpID & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetEmployeeBranch = IIf(IsNull(Rs3("BranchId").value), 0, Rs3("BranchId").value)
Else
GetEmployeeBranch = 0
End If
End Function
Private Sub DBCboClientName_Change()

    On Error Resume Next
    Dim lblflag  As Integer
     TxtCustCode.text = ""
    Dim DefaultSalesPersonId As Integer
    Dim fullcode As String
    Dim CurrncyID As Integer
    Dim Account_code As String
    DBCboClientName_Click (0)
    Me.DcboDebitSide.BoundText = DBCboClientName.BoundText
     Me.DcbCurrency.BoundText = MainCurrency()
If DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1 Or DCboCashType.ListIndex = 13 Or DCboCashType.ListIndex = 14 Then
    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode, , , , , , CurrncyID
    TxtCustCode.text = fullcode
    Me.DcbCurrency.BoundText = CurrncyID
ElseIf DCboCashType.ListIndex = 4 Then
Me.DcboDebitSide.BoundText = DBCboClientName.BoundText
If Option4.value = True Then
EmpIDD = GetEmpID("Account_Code1", DBCboClientName.BoundText)
End If
If Option5.value = True Then
EmpIDD = GetEmpID("Account_Code", DBCboClientName.BoundText)
End If
If Option6.value = True Then
EmpIDD = GetEmpID("Account_Code2", DBCboClientName.BoundText)
End If
If Option7.value = True Then
EmpIDD = GetEmpID("Account_Code3", DBCboClientName.BoundText)
End If
        If Option4.value = True Then
        lblflag = 1
       ElseIf Option5.value = True Then
        lblflag = 0

       ElseIf Option6.value = True Then
        lblflag = 2
      ElseIf Option7.value = True Then
        lblflag = 3
       End If

  GetEmployeeIDFromCode , , , fullcode, , lblflag, DBCboClientName.BoundText, True
       TxtCustCode.text = fullcode
       Me.DcbEmpBranch.BoundText = GetEmployeeBranch(EmpIDD)
   
   ElseIf val(Me.DCboCashType.ListIndex) = 5 Then
   TxtCustCode.text = getAccountSerial_Code("Account_Serial", "Account_Code", DBCboClientName.BoundText)
End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If DCboCashType.ListIndex = 3 Or DCboCashType.ListIndex = 4 Or DCboCashType.ListIndex = 5 Or DCboCashType.ListIndex = 6 Then
        If DCboCashType.ListIndex = 3 Then
       ReloadContrac val((DBCboClientName.BoundText))
          GetProjectsDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode
           TxtCustCode.text = fullcode
           Me.DcboDebitSide.BoundText = GetProjectCoount(val((DBCboClientName.BoundText)))
        End If
        
'            Me.DcboDebitSide.BoundText = GetProjectCoount(val((DBCboClientName.BoundText)))
        
        
        
        ElseIf DCboCashType.ListIndex = 2 Then
            If SystemOptions.SuppCreat4Acc = True Then
                If subContOpt(0).value = True Then
                    Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_CodeHi1")
                ElseIf subContOpt(1).value = True Then
                    Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_CodeAss2")
                ElseIf subContOpt(2).value = True Then
                    Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_Code")
                End If
            ElseIf SystemOptions.SubContactorHave3Account = False Then
            
            Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
            Else
            
           If subContOpt(2).value = True Then
           Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
           ElseIf subContOpt(1).value = True Then
           Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code1")
           ElseIf subContOpt(0).value = True Then
           Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code2")
           End If
            End If
            
     Else
          Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
          
        End If
    End If
If Me.TxtModFlg.text <> "R" Then
txtperson.text = DBCboClientName.text
End If
End Sub
Function GetProjectCoount(Optional ID As Double) As String
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = " SELECT  expanses_account ,AccountUnderImp  "
sql = sql & " From dbo.Projects"
sql = sql & " WHERE     (id =" & ID & ")"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetProjectCoount = IIf(IsNull(Rs4("expanses_account").value), IIf(IsNull(Rs4("AccountUnderImp").value), "", Rs4("AccountUnderImp").value), Rs4("expanses_account").value)
Else
GetProjectCoount = ""
End If
End Function
Private Sub DBCboClientName_Click(Area As Integer)
  If Me.TxtModFlg.text <> "R" Then
    If val(Me.CboPayMentType.ListIndex) <> -1 Or val(Me.CboPayMentType.ListIndex) <> 0 Then
    If val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 1 Or val(DCboCashType.ListIndex) = 2 Then
    GetCustomer val(DBCboClientName.BoundText)
    End If
    If val(Me.DCboCashType.ListIndex) = 4 Then
    GetEmployee EmpIDD
    End If
    End If
    End If
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If DCboCashType.ListIndex = 0 Then
        If KeyCode = vbKeyF3 Then
          FrmCustemerSearch.SearchType = 4
            FrmCustemerSearch.show vbModal
            
        End If

    ElseIf DCboCashType.ListIndex = 1 Or DCboCashType.ListIndex = 13 Then

        If KeyCode = vbKeyF3 Then
    
        FrmCompanySearch.lblSearchtype.Caption = 3
          FrmCompanySearch.show vbModal
       
        End If




    ElseIf DCboCashType.ListIndex = 2 Or DCboCashType.ListIndex = 14 Then

        If KeyCode = vbKeyF3 Then
   
     FrmCompanySearch.lblSearchtype.Caption = 9
       FrmCompanySearch.show vbModal
    
        End If
        
        
    ElseIf DCboCashType.ListIndex = 3 Then

        If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 5
           FrmProjectSearch.show vbModal
           
        End If
        
    ElseIf DCboCashType.ListIndex = 5 Then

        If KeyCode = vbKeyF3 Then
        DBCboClientName.text = ""
            Unload Account_search
            Account_search.show
           Account_search.case_id = 193
            
        End If

 ElseIf DCboCashType.ListIndex = 4 Then

    If KeyCode = vbKeyF3 Then

        FrmEmployeeSearch.lbltype = 34

              If Option4.value = True Then
             FrmEmployeeSearch.lblflag = 1
             ElseIf Option5.value = True Then
            FrmEmployeeSearch.lblflag = 0

             ElseIf Option6.value = True Then
            FrmEmployeeSearch.lblflag = 2
            ElseIf Option7.value = True Then
           FrmEmployeeSearch.lblflag = 3
             End If



       Set FrmEmployeeSearch.RetrunFrm = Me

       FrmEmployeeSearch.show
  
    End If





    End If

End Sub
Sub GetDataOfBank(Optional BankID As Double = 0)
If BankID <> 0 Then
Dim sql As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
sql = "select ReportName from BanksData where BankID=" & BankID & ""
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
TxtReportName.text = IIf(IsNull(Rs6("ReportName").value), "", Rs6("ReportName").value)
Else
TxtReportName.text = ""
End If
End If
End Sub
Private Sub DcboBankName_Click(Area As Integer)
    On Error Resume Next

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        '    Me.DcboCreditSide.BoundText = "a2a3a2"
        GetDataOfBank val(Me.DcboBankName.BoundText)
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If SystemOptions.banks_Accounts3 = True Then
            Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code2")
        Else
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
        
        If CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then
                     
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If
        
        'Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value

    End If

End Sub

Private Sub DcboBox_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub
Function getAccountSerial_Code(Optional filed As String, Optional FiledWher As String, Optional str As String) As String
Dim My_SQL As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If " & Filed &" <> "" Then
 My_SQL = "  select " & filed & " as Acoud from ACCOUNTS where " & FiledWher & "='" & str & "'"
 Rs7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
 If Rs7.RecordCount > 0 Then
 getAccountSerial_Code = IIf(IsNull(Rs7("Acoud").value), "", Rs7("Acoud").value)
 Else
 getAccountSerial_Code = ""
 End If
 End If
End Function
Private Sub DCboCashType_Change()
lbl(65).Visible = False
TxtPrePayd(17).Visible = False
lbl(70).Visible = False
DcbContractor.Visible = False
lbl(90).Visible = False
lbl(49).Visible = False
 
 TxtVATValue.Visible = False
 txtVat2.Visible = False
 lbl(98).Visible = False

DcbEmpBranch.Visible = False
TxtEndService.Visible = False
    Frame2.Visible = False
    Option4.value = False
    Option5.value = False
    Option6.value = False
    Option7.value = False
    Frame5.Visible = False
    Frame6.Visible = False
    Frame7.Visible = False
    Frame16.Visible = False
    Frame12(3).Visible = False
    Command5(0).Visible = False
    Command5(1).Visible = False
    Command5(2).Visible = False
    Frame9.Visible = False
    DBCboClientName.Visible = True
    lbl(3).Visible = True
    TxtCustCode.Visible = True
    XPTxtVal.Enabled = True
        Frame8.Visible = False
     Fra(4).Visible = False
    DBCboClientName = ""
     XPTxtVal.Enabled = True
   Fra(2).Visible = False
       lbl(47).Visible = False
        TxtAdvance.Visible = False
lbl(54).Visible = False
TxtDue.Visible = False
    
DBCboClientName.Visible = True

TxtCustCode.Visible = True

lbl(3).Visible = True
    Dim StrSQL As String
    Dim intDef As Integer
    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String

    If SystemOptions.UserInterface = EnglishInterface Then
        lbl(3).Caption = "Name"
    Else
        lbl(3).Caption = "«·«”„"
    End If
        
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "E" Then

        With FG
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows + 1
    
        End With

        Fra(2).Visible = False
        lbl(47).Visible = False
        TxtAdvance.Visible = False
    End If

    If DCboCashType.ListIndex = 13 Or DCboCashType.ListIndex = 14 Then
        TXT_order_no.Visible = True
        lbl(37).Visible = True
        lbl(37).Caption = "—ﬁ„ «·«⁄ „«œ"
        txtAcceptianPeriod.Visible = True
        lbl(101).Visible = True
    Else

        txtAcceptianPeriod.Visible = False
        lbl(101).Visible = False
    End If

    Select Case DCboCashType.ListIndex

        Case 0
        
       '     TxtVATValue.Visible = True
            txtVat2.Visible = True
            lbl(98).Visible = True
            
            Set Dcombos = New ClsDataCombos
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
            ChkTrans.Visible = True
            Fra(0).Visible = True

        Case 1, 13
            Set Dcombos = New ClsDataCombos
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
            ChkTrans.Visible = True
            Fra(0).Visible = True
            Command5(0).Visible = True
            Command5(1).Visible = True
            lbl(65).Visible = True
            TxtPrePayd(17).Visible = True
'            TxtVATValue.Visible = True
            txtVat2.Visible = True
            lbl(98).Visible = True
            
        Case 2, 14
        TxtPrePayd(17).Visible = True
        
            Set Dcombos = New ClsDataCombos
            Dcombos.GetPersons Me.DBCboClientName
            ChkTrans.Visible = False
            Fra(0).Visible = False
Frame5.Visible = True
subContOpt(2).value = True
'            TxtVATValue.Visible = True
            txtVat2.Visible = True
            lbl(98).Visible = True
            
            

        Case 3
             DcbContractor.Visible = True
            lbl(90).Visible = True
            Fra(0).Visible = True
            Command5(2).Visible = True
            If SystemOptions.UserInterface = EnglishInterface Then
                lbl(3).Caption = "Project"
            Else
                lbl(3).Caption = "«·„‘—Ê⁄"
            End If

            Set Dcombos = New ClsDataCombos
            Dcombos.ClearMyDataCombo Me.DBCboClientName
            If SystemOptions.UserInterface = ArabicInterface Then
                    My_SQL = "  select ID,Project_name from projects where not(Project_name is null) "
                     StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
                  StrSQL = StrSQL & "   order by Project_name " '  where  Account_code like'" & Account_Code_dynamic & "%' and last_account=1"
            Else
                    My_SQL = "  select ID,Project_namee from projects where not(Project_namee is null)"
                     StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
                    StrSQL = StrSQL & "  order by Project_name " '  where  Account_code like'" & Account_Code_dynamic & "%' and last_account=1"
            End If
            
            fill_combo Me.DBCboClientName, My_SQL
 
        Case 4
        
        lbl(49).Visible = True
        DcbEmpBranch.Visible = True
        Frame2.Visible = True
        
     If TxtOrder.text = "" Then
       Frame2.Enabled = True
  Else
    Frame2.Enabled = False
  End If
            If Me.TxtModFlg = "N" Then
  
                'DBCboClientName.DataSource = Nothing
                If SystemOptions.UserInterface = ArabicInterface Then
                         My_SQL = "  select Account_Code,Account_Name from ACCOUNTS where  Account_Code='0000'"
                 Else
                  My_SQL = "  select Account_Code,Account_Nameeng  from ACCOUNTS where  Account_Code='0000'"
                 End If
            
            Else
                        If SystemOptions.UserInterface = ArabicInterface Then
                            My_SQL = "  select Account_Code,Account_Name from ACCOUNTS where last_account=1"
                        Else
                        My_SQL = "  select Account_Code,Account_Nameeng from ACCOUNTS where last_account=1"
                        End If
            End If
       My_SQL = My_SQL & GetAccountByBarnchUser
            Set Dcombos = New ClsDataCombos
            Dcombos.ClearMyDataCombo Me.DBCboClientName
         
            fill_combo Me.DBCboClientName, My_SQL
      
        Case 5
TxtPrePayd(17).Visible = True

            Set Dcombos = New ClsDataCombos
            Dcombos.ClearMyDataCombo Me.DBCboClientName
If SystemOptions.UserInterface = ArabicInterface Then
            My_SQL = "  select Account_Code,Account_Name from ACCOUNTS where last_account=1"
Else
My_SQL = "  select Account_Code,Account_Nameeng from ACCOUNTS where last_account=1"
End If
 My_SQL = My_SQL & GetAccountByBarnchUser
            fill_combo Me.DBCboClientName, My_SQL
       
            '   My_SQL = "  select Account_Code,BoxName from TblBoxesData where Type=1"
            '  fill_combo Me.DBCboClientName, My_SQL
 
'            TxtVATValue.Visible = True
            txtVat2.Visible = True
            lbl(98).Visible = True

        Case 6
        XPTxtVal.Enabled = False
        DBCboClientName.Visible = False
       TxtCustCode.Visible = False
'   If SystemOptions.UserInterface = ArabicInterface Then
'            My_SQL = "  select Account_Code,BoxName from TblBoxesData where Type=0"
'    Else
'
'    My_SQL = "  select Account_Code,BoxNameE from TblBoxesData where Type=0"
'
'    End If
'
'            fill_combo Me.DBCboClientName, My_SQL
'
    
 'Case 7
 Frame6.Visible = True
 Case 7
    DBCboClientName.Visible = False
    lbl(3).Visible = False
    TxtCustCode.Visible = False
    XPTxtVal.Enabled = False
    lbl(65).Visible = True
    TxtPrePayd(17).Visible = True
    Frame7.Visible = True
    XPTxtVal.Enabled = False
'            TxtVATValue.Visible = True
            txtVat2.Visible = True
            lbl(98).Visible = True
    
 Case 8
  DBCboClientName.Visible = False
     lbl(3).Visible = False
    TxtCustCode.Visible = False
    XPTxtVal.Enabled = False
 lbl(54).Visible = True
TxtDue.Visible = True
   Frame15.Visible = False
 Frame8.Visible = True
 XPTxtVal.Enabled = False
 DBCboClientName.Visible = False

TxtCustCode.Visible = False

lbl(3).Visible = False
 Case 9
 Fra(4).Visible = True
 DBCboClientName.Visible = False
     lbl(3).Visible = False
    TxtCustCode.Visible = False
    XPTxtVal.Enabled = False
 Case 10
 Frame12(3).Visible = True
  DBCboClientName.Visible = False
     lbl(3).Visible = False
    TxtCustCode.Visible = False
    XPTxtVal.Enabled = False
    lbl(70).Visible = True
    lbl(70).Caption = "”‰œ ‰Â«Ì… Œœ„…"
TxtEndService.Visible = True
Frame15.Visible = False
Case 11
  Frame16.Visible = True
  DBCboClientName.Visible = False
  lbl(3).Visible = False
  TxtCustCode.Visible = False
  XPTxtVal.Enabled = False
  XPTxtVal.Enabled = False
  Case 12
  DBCboClientName.Visible = False
  XPTxtVal.Enabled = False
      lbl(70).Visible = True
    lbl(70).Caption = "—ﬁ„ «·«ﬁ—«— «·÷—Ì»Ì"
TxtEndService.Visible = True
TxtEndService.Enabled = True
TxtEndService.locked = False
DcboDebitSide.BoundText = get_account_code_branch(145, my_branch)
    End Select
CalCulteVAT
    cSearchDcbo.Refresh
    Set Dcombos = Nothing
    Exit Sub
ErrTrap:

End Sub

Private Sub DCboCashType_Click()
    DCboCashType_Change
End Sub
Sub CalCulteVATOld()
Dim AccountVATCreit As String
Dim Percetage As Double
If Me.TxtModFlg.text <> "R" Then
If Option3.value = True And val(DCboCashType.ListIndex) = 1 Then
          GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
          PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
          If XPTxtVal.text <> "" Then
      TxtPrePayd(17).text = Format(XPTxtVal.text, "###.00") * Percetage / 100
      End If
Else
'TxtPrePayd(17).Text = 0
End If
End If
End Sub

Public Sub CalCulteVAT(Optional Ind As Integer = 0)
Dim AccountVATCreit As String
Dim Percetage As Double
Dim mDigit As Integer

        If Option3.value Then
           mDigit = 1
        Else
            mDigit = val(SystemOptions.SysDefCurrencyForamt)
            
        End If
If Option3.value = True And (val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 1 Or val(DCboCashType.ListIndex) = 2 Or val(DCboCashType.ListIndex) = 3 Or val(DCboCashType.ListIndex) = 5 Or val(DCboCashType.ListIndex) = 7) Then

    If Ind = 3 Then
        PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
           TxtPrePayd(17).text = val(Format(val(val(Format((XPTxtVal.text), "###.00")) * Percetage / 100), "." & String(Abs(mDigit), "#")))
           'val(Format((XPTxtVal.Text), "###.00")) * Percetage / 100
         TxtVATValue.text = val(Format(val(val(Format((XPTxtVal.text), "###.00")) * Percetage / 100), "." & String(Abs(mDigit), "#")))
          
         txtVat2.text = TxtVATValue.text
         Exit Sub
    End If
    'XPDtbTrans.value = 100
    'XPTxtVal = 100
    
    If Me.TxtModFlg.text <> "R" And Me.TxtModFlg.text <> "" Then
    
    CalcTotal Ind
    If Option3.value = True And (val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 1 Or val(DCboCashType.ListIndex) = 2 Or val(DCboCashType.ListIndex) = 3 Or val(DCboCashType.ListIndex) = 5 Or val(DCboCashType.ListIndex) = 7) Then
              
    If SystemOptions.NotAllowedCalcVata Then
        TxtVATValue.text = 0
        txtVat2.text = 0
        TxtPrePayd(17) = 0
    Else
        GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
             
        PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
    
         TxtPrePayd(17).text = val(Format(val(val(Format((XPTxtVal.text), "###.00")) * Percetage / 100), "." & String(Abs(mDigit), "#")))
        ' val (Format((XPTxtVal.Text), "###.00")) * Percetage / 100
         Dim mVal As Double
         mVal = val(Format((XPTxtVal.text), "###.00"))
         TxtVATValue.text = val(Format(val(val(Format((XPTxtVal.text), "###.00")) * Percetage / 100), "." & String(Abs(mDigit), "#")))
        ' val (Format((mVal), "###.00")) * Percetage / 100
         txtTotalWithVat.text = Round(val(Format((mVal), "###.00")) + val(TxtVATValue.text), 2)
            
    End If
          
    Else
    TxtVATValue.text = 0
    TxtPrePayd(17) = 0
    End If
    
    End If
    
    txtVat2.text = TxtVATValue.text
Else
 '   TxtPrePayd(17) = 0
 txtVat2.text = 0
    TxtVATValue.text = 0
    txtTotalWithVat.text = myRound(XPTxtVal.text)
    TxtPrePayd(17).text = 0
End If
txtVat2.text = TxtVATValue.text
End Sub
Sub CalcTotal(Optional Ind As Integer)
     Dim mDigit As Integer
     If Option3.value Then
           mDigit = 1
        Else
            mDigit = val(SystemOptions.SysDefCurrencyForamt)
            
        End If
    If Ind = 1 Then
    
        txtTotalWithVat = Round(val(txtVat2), mDigit) + myRound(XPTxtVal)
    ElseIf Ind = 0 Then
           Dim Percetage As Double
    Dim AccountVATCreit As String
     
    If SystemOptions.NotAllowedCalcVata Then
        TxtVATValue.text = 0
        txtVat2.text = 0
        TxtPrePayd(17) = 0
    Else
        PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
    End If
     
    
    'TxtVATValue.Text = val(XPTxtVal.Text) * Percetage / 100
    If Option3.value = True Then
    XPTxtVal.text = Round(val(txtTotalWithVat) / (Percetage / 100 + 1), 2)
    
    TxtVATValue.text = Round(myRound(XPTxtVal.text) * Percetage / 100, mDigit)
    
    
    txtVat2.text = TxtVATValue.text
    Else
      XPTxtVal.text = val(txtTotalWithVat.text)
    End If
    End If
End Sub
'Private Function myRound(ByVal mTxt As String, Optional ByVal mR As Integer = 0) As Double
'    If mR = 0 Then mR = 2
'    If Trim(mTxt) = "" Then
'        myRound = val(mTxt)
'    Else
'        myRound = Round(val(CDbl(mTxt)), mR)
'    End If
'
'End Function

Private Function myRound(ByVal mNumber As Variant, _
                        Optional NoOfDecimalDigits As Integer) As Double
    Dim X As Double

    If IsNumeric(Trim(mNumber)) Then X = CDbl(Trim(mNumber)) Else X = val(Trim(mNumber))
    '-------------------------
    If X = 0 Then myRound = 0 Else myRound = Round(X + 1E-17, IIf(NoOfDecimalDigits = 0, 2, NoOfDecimalDigits))
End Function


Sub ClaCul()

    'Me.lbl(18).Caption = WriteNo(Me.XPTxtVal.text, 0, True)
    'txtAdv_payment_value.text = Format(Val(XPTxtVal.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
'calcnet
    If SystemOptions.NotAllowedCalcVata Then
        TxtVATValue.text = 0
        txtVat2.text = 0
        TxtPrePayd(17) = 0
    End If
    CalCulteVAT 1
    If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(18).Caption = WriteNo(Format(val(XPTxtVal.text) + val(TxtVATValue.text), "0.00"), 0, True, ".", , 0)

    Else
 
        Me.lbl(18).Caption = WriteNo(Format(val(XPTxtVal.text) + val(TxtVATValue.text), "0.00"), 0, True, ".", , 1)

    End If

    'If TxtModFlg.text = "N" Or TxtModFlg.text = "E" And Option3.value = True Then
    If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
        txtAdv_payment_value.text = XPTxtVal.text
    End If

   If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(18).Caption = WriteNo(Format(val(XPTxtVal.text) + val(TxtVATValue.text), "0.00"), 0, True, ".", , 0)

    Else
 
        Me.lbl(18).Caption = WriteNo(Format(val(XPTxtVal.text) + val(TxtVATValue.text), "0.00"), 0, True, ".", , 1)

    End If
End Sub

Private Sub ChangeLang()


    TranslateForm Me, True
XPLbl(15).Caption = "Salary"
Check2.RightToLeft = False
Option8(0).RightToLeft = False
Option8(1).RightToLeft = False
Option8(0).Caption = "Purchases"
Option8(1).Caption = "Financial"
Check2.Caption = "Select All"
Label27(2).Caption = "Total"
Frame16.Caption = "Allowances"
lbl(89).Caption = "Bank Expenses"
XPLbl(25).Caption = "Other Discount"
Command5(2).Caption = "Show Project Bills"
Command12.Caption = "Show"
IncludVAT.Caption = "Include VAT"
lbl(65).Caption = "VAT"
lbl(90).Caption = "Contractor"
lbl(91).Caption = "Account"
lbl(92).Caption = "Rate"
lbl(96).Caption = "Value"
lbl(93).Caption = "Currency"
 XPLbl(28).Caption = "Prepaid"
 XPLbl(30).Caption = "Prepaid"
 XPLbl(31).Caption = "Prepaid"
 XPLbl(32).Caption = "Prepaid"
 XPLbl(33).Caption = "Prepaid"
 XPLbl(34).Caption = "Prepaid"
 XPLbl(35).Caption = "Prepaid"
 XPLbl(36).Caption = "Prepaid"
 XPLbl(37).Caption = "Prepaid"
 XPLbl(38).Caption = "Prepaid"
 XPLbl(39).Caption = "Prepaid"
    lbl(22).Caption = "Curr. Week"
    lbl(40).Caption = "Branch"
    Frame12(3).Caption = "Data of End Service"
    XPLbl(17).Caption = "Ticket"
    Label22.Caption = "Employee"
    Label21.Caption = "Branch"
    XPLbl(14).Caption = "Net Paid"
    XPLbl(26).Caption = "End Service"
    XPLbl(23).Caption = "Total Discount"
    XPLbl(21).Caption = "Add"
    XPLbl(16).Caption = "Custom Tickets"
    lbl(88).Caption = "Prepaid"
    lbl(87).Caption = "Prepaid"
    lbl(86).Caption = "Prepaid"
    lbl(85).Caption = "Prepaid"
    lbl(84).Caption = "Prepaid"
    lbl(83).Caption = "Prepaid"
    lbl(82).Caption = "Prepaid"
    XPLbl(27).Caption = "Paid"
     lbl(75).Caption = "Paid"
    lbl(76).Caption = "Paid"
    lbl(77).Caption = "Paid"
    lbl(78).Caption = "Paid"
    lbl(79).Caption = "Paid"
    lbl(80).Caption = "Paid"
    lbl(81).Caption = "Paid"
    XPLbl(0).Caption = "Paid"
    XPLbl(1).Caption = "Paid"
    XPLbl(4).Caption = "Paid"
    XPLbl(6).Caption = "Paid"
    'XPLbl(15).Caption = "Paid"
    XPLbl(20).Caption = "Paid"
    XPLbl(2).Caption = "Paid"
    XPLbl(24).Caption = "Paid"
    XPLbl(8).Caption = "Paid"
    XPLbl(3).Caption = "Paid"
    XPLbl(9).Caption = "Paid"
    XPLbl(22).Caption = "Paid"
    XPLbl(40).Caption = "Paid"
    XPLbl(29).Caption = "Tickets from Cont."
      XPLbl(19).Caption = "Curr Salary"
    XPLbl(17).Caption = "Ticket Value"
    XPLbl(18).Caption = "Vacation Value"
    lbl(70).Caption = "End Service"
    lbl(73).Caption = "Branch"
 XPLbl(7).Caption = "ReSult"
XPLbl(10).Caption = "Advance"
XPLbl(13).Caption = " Without Pay"
XPLbl(11).Caption = "Testament"
Command1.Caption = "Show"
Command4.Caption = "Show"
Command8.Caption = "Show"
lbl(72).Caption = "Year"
lbl(74).Caption = "Insur.Value"
lbl(71).Caption = "Month"
XPLbl(5).Caption = "Total"
XPLbl(12).Caption = "Net"
Command11.Caption = "Cancel Payment"
Command3.Caption = "Cancel Payment"
Command2.Caption = "Cancel Payment"
Command7.Caption = "Cancel Payment"
Command6.Caption = "Cancel Payment"
Command10(0).Caption = "Cancel Payment"
Command10(1).Caption = "Cancel Payment"
Check18.RightToLeft = False
Check18.Caption = "Select All"
Check1.RightToLeft = False
Check1.Caption = "Select All"
Label27(0).Caption = "Total"
lbl(58).Caption = "Name"
lbl(56).Caption = "ID"
lbl(57).Caption = "Phone"
'''////////
lbl(97).Caption = "Address"
lbl(59).Caption = "ID"
'lbl(65).Caption = "Date"
lbl(67).Caption = "Issued At"
lbl(68).Caption = "Birth At"
lbl(60).Caption = "Cuntry/Gover"
Label26.Caption = "Account No"
Label19.Caption = "Banck"
Label20.Caption = "IBAN"
Label24.Caption = "Bank Code"
Label23.Caption = "Banck Address"
lbl(69).Caption = "City/Sreet"
Label18.Caption = "Phone"
Cmd(13).Caption = "Print Deposit "
Frame5.Caption = "Subcontractor"
Command5(1).Caption = "Show Bills Purchases "
Frame12(0).Caption = "Data of Bills Purchases "
Frame12(1).Caption = "Data of Project Bills"

Command5(0).Caption = "Show Finance Bills"

Frame9.Caption = "Finance Bills Data"
''///
subContOpt(0).RightToLeft = False
subContOpt(1).RightToLeft = False
subContOpt(2).RightToLeft = False
subContOpt(0).Caption = "Prepayment"
subContOpt(1).Caption = "Warranty"
subContOpt(2).Caption = "Works"
''///
Frame8.Caption = "Vacation Entitlement Data"
Label13.Caption = "Employee"
Label14.Caption = "Branch"
Label5.Caption = "Salary this Month"
Label6.Caption = "Other Benefits"
Label7.Caption = "Other Deductions "
Label9.Caption = "Previous Loans"
Label10.Caption = "Total Entitlement Payment"
Label12.Caption = "Vac. Entitlement"
Label11.Caption = "Tickets"
'''//
lbl(34).Caption = "Name"
lbl(61).Caption = "Country/Gove."
lbl(64).Caption = "Guar./Phone"
lbl(66).Caption = "Guar.Address"

Frame7.Caption = "Prepayments"
Fra(4).Caption = "Request Exchange"
lbl(63).Caption = "City/Str"
lbl(62).Caption = "Address"
Frame11.Caption = "Transfer Data"
lbl(55).Caption = "Req .No"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Cmd(12).Caption = "Same Copy"
    Label3.Caption = "Month"
    Label4.Caption = "Year"
    Fra(3).Caption = "Options"
    Frame6.Caption = "Salaries"
    lbl(48).Caption = "Project"
    lbl(50).Caption = "Pand"
    lbl(51).Caption = "Process"
    lbl(53).Caption = "Manual. No"
    Fra(2).Caption = "detalis"
    Cmd(15).Caption = "Attachments"
lbl(46).Caption = "Order Exchane"
lbl(47).Caption = "Order Advance"
    lbl(52).Caption = "Management"
    Option3.Caption = "Adv. Payment"
    Option2.Caption = "Select Invoice"
    ALLButton3.Caption = "Select"
    lbl(37).Caption = "Order No :"
    lbl(22).Caption = "Current Week"
    lbl(35).Caption = "Adv. Pay."
    lbl(94).Caption = "General C.C."
    lbl(36).Caption = "General Des"
    Cmd(9).Caption = "GL Print"
    Cmd(10).Caption = "Cheque Print"
    Frame2.Caption = "Employee"
    Option4.Caption = "Salary"
    Option5.Caption = "Advanced"
    Option6.Caption = "Alloc"
    Option7.Caption = "Adv. Paayment"

   ' ALLButton1.Caption = "Installment view"
   ' ALLButton2.Caption = "debt Voucher"
    Me.Caption = "Payable Voucher"
    C1Elastic1.Caption = Me.Caption
    lbl(4).Caption = "Opr Code"
    lbl(1).Caption = "Date"
    lbl(0).Caption = "Type"
    lbl(3).Caption = "Name"
    lbl(2).Caption = "Value"
    lbl(14).Caption = "Payemnt Method"
    lbl(9).Caption = "Box Name"
    lbl(15).Caption = "Bank Name"
    lbl(16).Caption = "Cheque #"
    lbl(17).Caption = "Cheque date"
    lbl(34).Caption = "Due To"
    lbl(5).Caption = "Note"
    ChkTrans.Caption = "From bill"
    lbl(12).Caption = "Bill type"
    lbl(10).Caption = "Bill #"
    lbl(13).Caption = "Current Balance"
    FraInfo.Caption = "Information"
    lbl(22).Caption = "Current Week"

    lbl(23).Caption = "Today Payments "
    lbl(27).Caption = "Cash"
    lbl(28).Caption = "Cheque"

    lbl(19).Caption = "Week Payments "

    lbl(21).Caption = "Cash"
    lbl(24).Caption = "Cheque"

    lbl(20).Caption = "Month Payments "

    lbl(25).Caption = "Cash"
    lbl(26).Caption = "Cheque"
    Fra(1).Caption = "GL"

    lbl(30).Caption = "GL#"
    lbl(29).Caption = "Interval"

    lbl(32).Caption = "Depit"
    lbl(31).Caption = "Credit"
    Cmd(8).Caption = "Table view"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Current "
    lbl(6).Caption = "Records Count "
    Frame4.Caption = "Bank Transfer Expenses"
    Label2.Caption = "Value"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    Cmd(14).Caption = "Help"
    DCboCashType.Clear
    DCboCashType.AddItem "To Customer"
    DCboCashType.AddItem "To Vendor"
    DCboCashType.AddItem "sub-contractor"
    DCboCashType.AddItem "To Project"
    DCboCashType.AddItem "To Employee"
    DCboCashType.AddItem "To Acc."
    DCboCashType.AddItem "Salaries"
    DCboCashType.AddItem "Prepayments"
    DCboCashType.AddItem "Vacation Due"
    DCboCashType.AddItem "To Suppller."
    DCboCashType.AddItem "End Service"
    DCboCashType.AddItem "Allowances"
    'DCboCashType.AddItem "Bety Cash"
    'DCboCashType.AddItem "Box Recharge"

    Option4.Caption = "Salary"
    Option5.Caption = "Advance"
    Option5.Caption = "Alloc"
    Option5.Caption = "Adv. Payment"

    With Me.CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
        .AddItem "Bank Transfer"
        .AddItem "P Cheque"
        .AddItem "Account"
        .AddItem "Credit payment"
      
    End With
    Label1.Caption = "You Can Manually Update Value Manual"
    Cmd(11).Caption = "Calculate"
    Label15.Caption = "Total"
    lbl(54).Caption = "Due Voucher"
    lbl(43).Caption = "First Date"
    ''///////
    Fra(2).Caption = "Payment Methods"
    lbl(44).Caption = "Payment No"
    lbl(42).Caption = "Month "
    lbl(41).Caption = "Year"
  With FG
  .TextMatrix(0, .ColIndex("PartNO")) = "No"
  .TextMatrix(0, .ColIndex("PartValue")) = "Value"
  .TextMatrix(0, .ColIndex("PartDate")) = "Date"
  End With
    ''//////
With VSFlexGrid2
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("payed")) = "Select"
.TextMatrix(0, .ColIndex("NoteSerial1")) = "Bill No"
.TextMatrix(0, .ColIndex("too")) = "Manual No."
.TextMatrix(0, .ColIndex("NoteDate")) = "Date"
.TextMatrix(0, .ColIndex("branch_name")) = "Branch"
.TextMatrix(0, .ColIndex("Note_Value")) = "Original value"
.TextMatrix(0, .ColIndex("PayedValue")) = "Payed Value"
.TextMatrix(0, .ColIndex("RemainingValue")) = "Remaining"
.TextMatrix(0, .ColIndex("TransPayedValue")) = "Payed Trans"
.TextMatrix(0, .ColIndex("NetValue")) = "Net Value"
.TextMatrix(0, .ColIndex("Project_name")) = "Project"
End With
    With VSFlexGrid1

.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("InstalValue")) = "Installment Value"
.TextMatrix(0, .ColIndex("haveqest")) = "Have Installments"
.TextMatrix(0, .ColIndex("payed")) = "Select"
.TextMatrix(0, .ColIndex("NoteSerial1")) = "Bill No"
.TextMatrix(0, .ColIndex("too")) = "Bill Supplier"
.TextMatrix(0, .ColIndex("NoteDate")) = "Date"
.TextMatrix(0, .ColIndex("branch_name")) = "Branch"
.TextMatrix(0, .ColIndex("Note_Value")) = "Original value"
.TextMatrix(0, .ColIndex("PayedValue")) = "Payed Value"
.TextMatrix(0, .ColIndex("RemainingValue")) = "Remaining"
.TextMatrix(0, .ColIndex("TransPayedValue")) = "Payed Trans"
.TextMatrix(0, .ColIndex("NetValue")) = "Net Value"
.TextMatrix(0, .ColIndex("Show")) = "Show"
.TextMatrix(0, .ColIndex("DueDate")) = "Due Date"

End With
With GRID1

.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("InstalValue")) = "Installment Value"
.TextMatrix(0, .ColIndex("haveqest")) = "Have Installments"
.TextMatrix(0, .ColIndex("payed")) = "Select"
.TextMatrix(0, .ColIndex("NoteSerial1")) = "Bill No"
.TextMatrix(0, .ColIndex("too")) = "Bill Supplier"
.TextMatrix(0, .ColIndex("NoteDate")) = "Date"
.TextMatrix(0, .ColIndex("branch_name")) = "Branch"
.TextMatrix(0, .ColIndex("Note_Value")) = "Original value"
.TextMatrix(0, .ColIndex("PayedValue")) = "Payed Value"
.TextMatrix(0, .ColIndex("RemainingValue")) = "Remaining"
.TextMatrix(0, .ColIndex("TransPayedValue")) = "Payed Trans"
.TextMatrix(0, .ColIndex("NetValue")) = "Net value"
.TextMatrix(0, .ColIndex("Show")) = "Show"

End With
lbl(49).Caption = "Emp.Branch"
    With Me.CboTrans
        .Clear
        .AddItem "Purchase invoice"
        .AddItem "Returned sales"
    End With

End Sub

Private Sub DcboDebitSide_Change()
    WriteCustomerBalPublic Me.DcboDebitSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
End Sub

Private Sub DcboEmpName_Change()
      If val(DcboEmpName.BoundText) = 0 Then Exit Sub


    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
    
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 5
    End If

End Sub

Private Sub dcopr_Click(Area As Integer)
Dcterm1_Change
End Sub

Private Sub DCPreFix_Change()
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg = "E" Then
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
  End If
End Sub

Private Sub DCPreFix_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    
End Sub

Private Sub dcproject1_Change()
If Me.TxtModFlg <> "R" Then
If val(dcproject1.BoundText) <> 0 Then
fillterms val(dcproject1.BoundText)
End If
End If

End Sub

Private Sub dcproject1_Click(Area As Integer)
dcproject1_Change
End Sub
Function fillterms(project_id As Integer)
    Dim My_SQL As String
 
    My_SQL = " select oprid,des from dbo.projects_des where project_id=" & project_id

    fill_combo Me.Dcterm1, My_SQL
       
        
    Dcterm1.ReFill
End Function



Private Sub Dcterm1_Change()
If Me.TxtModFlg <> "R" Then
 Dim Dcombos As ClsDataCombos

       Set Dcombos = New ClsDataCombos
  If dcproject1.BoundText <> "" Then
        
         If Me.Dcterm1.BoundText <> "" Then
         Dcombos.GetProcessOfProjedt dcopr, val(dcproject1.BoundText), , val(Dcterm1.BoundText), 2
         End If
       
    End If
 End If
 
End Sub

Private Sub Dcterm1_Click(Area As Integer)
'dcproject1_Change
'fillterms val(DCPROJECT1.BoundText)

'Dcterm1_Change
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Me.TxtModFlg.text <> "R" Then
With FG
.TextMatrix(Row, .ColIndex("PartValue")) = Abs(val(.TextMatrix(Row, .ColIndex("PartValue"))))

End With
End If
Me.LblTotalV.Caption = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("PartValue"), FG.rows - 1, FG.ColIndex("PartValue"))
End Sub

Private Sub Form_Load()
Dim My_SQL As String

           If SystemOptions.MonyeIssueVchrNoMust = True Then
           TxtOrder.locked = True
              End If
           
              
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.DateOpt = 1 Then
        Txt_DateHigri.Visible = True
    
    End If

                   If SystemOptions.SpecialVersion = True Then
Cmd(9).Visible = False
Fra(1).Visible = False
   End If
   
 On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos
     My_SQL = " select id,code from currency"
    fill_combo Me.DcbCurrency, My_SQL
    'My_SQL = " select id,Project_name from projects order by Project_name"
    
        If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "SELECT id,Project_name From  projects "
        My_SQL = My_SQL & " where  not (Project_name is null)and Project_name<>N'""'"
    Else
        My_SQL = "SELECT id,Project_nameE From   projects "
        My_SQL = My_SQL & " where  not (Project_nameE is null)and Project_nameE<>N'""'"
    End If
    My_SQL = My_SQL & " and (Not (Fullcode Is Null))"
    If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = My_SQL & " order by  Project_name"
    Else
    My_SQL = My_SQL & " order by  Project_nameE"
    End If
    
    
    fill_combo dcproject1, My_SQL
    My_SQL = " select  oprid,des from projects_des"
    fill_combo Dcterm1, My_SQL

    My_SQL = " select  id,name from terms_operations"
    fill_combo dcopr, My_SQL

    My_SQL = " Select id , name from  TblDurations "
    fill_combo dcDur, My_SQL

ReloadContracR

    ScreenNameArabic = "”‰œ ’—› - «·„œ›Ê⁄«   "
    ScreenNameEnglish = "Payable Voucher"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 5
 
    YearMonth
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set Cmd(14).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
'    Resize_Form Me

    AddTip
    DCboCashType.AddItem "≈·Ï ⁄„Ì·"
    DCboCashType.AddItem "≈·Ï „Ê—œ"
    DCboCashType.AddItem "„ﬁ«Ê· »«ÿ‰"
    DCboCashType.AddItem "„‘—Ê⁄"
    DCboCashType.AddItem "„ÊŸ›"
    DCboCashType.AddItem "Õ”«»"
    DCboCashType.AddItem "—Ê« »"
    DCboCashType.AddItem "„œ›Ê⁄«  „ﬁœ„…"
    DCboCashType.AddItem " „” Õﬁ«  ≈Ã«“…"
    DCboCashType.AddItem "”‰œ ’—› „ ⁄ÂœÌ‰"
    DCboCashType.AddItem "‰Â«Ì… Œœ„…"
    DCboCashType.AddItem "»œ·«  „ﬁœ„…"
    DCboCashType.AddItem "„œ›Ê⁄«  «·«ﬁ—«— «·÷—Ì»Ì"
    DCboCashType.AddItem "«·«⁄ „«œ«  «·„ﬁ»Ê·… ·„Ê—œ"
    DCboCashType.AddItem "«·«⁄ „«œ«  «·„ﬁ»Ê·… ·„ﬁ«Ê· »«ÿ‰"
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmpDepartments DcbDepartment
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetBranches Me.DcbBrReq
    Dcombos.GetAccountingCodes Me.DcbAccount, True, False
    Dcombos.GetBranches Me.dcBranch1
    Dcombos.GetEmployees Me.DcbEmpEndService
    Dcombos.GetBranches Me.DcbBranchEndServ
    Dcombos.GetCostCenter DcCostCenter
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetPrefix2 Me.DCPreFix, 4, 0
    Dcombos.GetBranches Me.DcbEmpBranch
   
    If SystemOptions.usertype <> UserAdmin Then
        Me.dcBranch.Enabled = True
    End If

    With Me.CboPayMentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "ÕÊ«·Â »‰ﬂÌÂ"
        .AddItem "  ‘Ìﬂ „”œœ"
        .AddItem "Õ”«»"
        .AddItem "¬Ã·"
    End With

    With Me.CboTrans
        .Clear
        .AddItem "›« Ê—… „‘ —Ì« "
        .AddItem "„— Ã⁄ „»Ì⁄« "
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    'Set cSearchDcbo = New clsDCboSearch
    'Set cSearchDcbo.Client = Me.DBCboClientName
    'cSearchDcbo(0).SetBuddyText Me.TxtCusID

    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide

    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=5 and ( cashingtype<=12 ) "
        
      
        
        If SystemOptions.usertype <> UserAdminAll Then
        StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
    
    StrSQL = "select * From Notes where NoteType=5 and ( cashingtype<=12 )     AND branch_no in(" & Current_branchSql & ")"
    StrSQL = StrSQL & " and   (  (akarid is null )  and   (IqarID2 is null )  and   (NoteOrBonID is null ) )  "
 
  If SystemOptions.FixedCustomer = 1 Then
                              StrSQL = StrSQL & " and  UserID = " & user_id
                               End If
                               
    StrSQL = StrSQL & "order by NoteID "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    SetDtpickerDate XPDtbTrans
    SetDtpickerDate Me.DtpChequeDueDate
    ChkTrans.value = Unchecked
    ChkTrans_Click

    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"
    WriteInfo
   

    'My_SQL = "  select account_no,account_name from projects  where not (account_no is null)"
    My_SQL = "  select expanses_account,Project_name from projects where not(expanses_account is null)" '  where  Account_code like'" & Account_Code_dynamic & "%' and last_account=1"
    fill_combo dcproject, My_SQL

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub
Sub ReloadContrac(Optional project_no As Double)
If Me.TxtModFlg.text <> "R" Then
Dim Dcombos As ClsDataCombos
Dim StrSQL As String
Set Dcombos = New ClsDataCombos
If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select CusID,CusName From TblCustemers"
Else
    StrSQL = "Select CusID,CusNamee From TblCustemers"
End If
'StrSQL = StrSQL & " where CusID in(SELECT     sub_contractor_id"
StrSQL = StrSQL & " where CusID in(SELECT     subContractorId"

'StrSQL = StrSQL & " From dbo.projects_des"
StrSQL = StrSQL & " From dbo.project_billl"

StrSQL = StrSQL & " WHERE     (project_no = " & project_no & "))"
Dcombos.ClearMyDataCombo DcbContractor
fill_combo Me.DcbContractor, StrSQL
ProjectIDD = project_no
End If
End Sub
Sub ReloadContracR()
Dim Dcombos As ClsDataCombos
Dim StrSQL As String
Set Dcombos = New ClsDataCombos
If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select CusID,CusName From TblCustemers"
Else
    StrSQL = "Select CusID,CusNamee From TblCustemers"
End If
'Dcombos.ClearMyDataCombo DcbContractor
fill_combo Me.DcbContractor, StrSQL

End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, 5

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

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Reline22
Reline2
Reline

End Sub

Private Sub Grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With GRID1
Select Case .ColKey(Col)
Case "TransPayedValue"
If Aut_manual = False Then

Cancel = True
Else
If .cell(flexcpChecked, Row, .ColIndex("payed")) = flexChecked Then
Cancel = False
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·›« Ê—… «Ê·«"
Else
MsgBox "Please Select Bill"
End If
End If
End If
Case "haveqest"
Cancel = True
Case "NoteSerial1"
Cancel = True
Case "too"
Cancel = True
Case "NoteDate"
Cancel = True
Case "branch_name"
Cancel = True
Case "Note_Value"
Cancel = True
Case "PayedValue"
Cancel = True
Case "RemainingValue"
Cancel = True
Case "NetValue"
Cancel = True
Case "InstalValue"
Cancel = True
End Select
End With
End Sub

Private Sub Grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With GRID1
Select Case .ColKey(Col)
Case "Show"
If val(.TextMatrix(Row, .ColIndex("NoteID"))) <> 0 And .cell(flexcpChecked, Row, .ColIndex("payed")) = flexChecked Then
'Unload FrmEmpSalary6
Load FrmEmpSalary6
FrmEmpSalary6.show
FrmEmpSalary6.ALLButton3.Visible = False
FrmEmpSalary6.ALLButton6.Visible = False

FrmEmpSalary6.Grid2.Visible = False
FrmEmpSalary6.GRID1.Visible = False
FrmEmpSalary6.Check17.Visible = False
FrmEmpSalary6.lbl(12).Visible = False
FrmEmpSalary6.DTPicker1.Visible = False
FrmEmpSalary6.VSFlexGrid1.Visible = False
FrmEmpSalary6.Check18.Visible = False
FrmEmpSalary6.Check19.Visible = False
FrmEmpSalary6.ALLButton7.Visible = False
FrmEmpSalary6.Check20.Visible = False
FrmEmpSalary6.VSFlexGrid2.Visible = False
FrmEmpSalary6.PayDes = .TextMatrix(Row, .ColIndex("StrQest"))
FrmEmpSalary6.ALLButton8.Visible = True
FrmEmpSalary6.VSFlexGrid3.Visible = True
FrmEmpSalary6.VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
FrmEmpSalary6.Check21.Visible = True
FrmEmpSalary6.FillGrid6 val(.TextMatrix(Row, .ColIndex("NoteID")))
FrmEmpSalary6.Row1 = Row
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— —ﬁ„   «·›« Ê—…"
Else
MsgBox "Please Select Bill  "
End If

Exit Sub
End If
End Select

End With
End Sub

Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With GRID1
Select Case .ColKey(Col)
Case "Show"
.ColComboList(.ColIndex("Show")) = "..."
End Select
End With
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption



End Sub

Private Sub Label17_Click()
Frame9.Visible = False
End Sub

Private Sub Label28_Click()
Frame12(1).Visible = False
End Sub

Private Sub Label29_Click()
Frame12(0).Visible = False
End Sub

Private Sub Label38_Click()
Frame15.Visible = False
End Sub

Private Sub Label39_Click()
Frame12(3).Visible = False
End Sub

Private Sub Label40_Click()
Frame8.Visible = False
End Sub

Private Sub LblLink_Click()
  If SystemOptions.SpecialVersion = True Then
        Exit Sub
End If
        
    'Dim LngCusID As Long
    'If DoPremis(Do_Print, "ReportCustomers", True) = False Then
    '    Exit Sub
    'End If
    'LngCusID = Val(Me.DBCboClientName.BoundText)
    'OpenScreen PopUpShowCustomerBalanceScreen, LngCusID, 0

    Dim FirstPeriod As Date
    getFirstPeriodDateInthisYear FirstPeriod
    ShowReport DcboDebitSide.BoundText, DcboDebitSide.text, FirstPeriod, Date

End Sub

Private Sub LblLink_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

    If SystemOptions.UserInterface = ArabicInterface Then
        LblLink.ToolTipText = "—’Ìœ «·ÿ—› «·„œÌ‰:" & WriteNo(Balance, 0, True)
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        LblLink.ToolTipText = "Depit Balance:" & WriteNo(Balance, 0, True)
    End If

End Sub



Private Sub numbering_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Option1_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If
 If Option1.value = True Then
Frame18.Visible = True
Else
Frame18.Visible = False
End If
CalCulteVAT
End Sub
Function AutoCalculate2() As Boolean
Dim i As Integer
Dim NetValu As Double
Dim TempValu As Double
Dim RemainValu As Double
NetValu = val(XPTxtVal.text)
With GRID1
For i = 1 To .rows - 1
RemainValu = val(.TextMatrix(i, .ColIndex("RemainingValue")))
If NetValu > RemainValu Then
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
AutoCalculate2 = False
Else
AutoCalculate2 = True
End If
End Function
Function AutoCalculate() As Boolean
Dim i As Integer
Dim NetValu As Double
Dim TempValu As Double
Dim RemainValu As Double
NetValu = val(XPTxtVal.text)
With VSFlexGrid1
For i = 1 To .rows - 1
RemainValu = val(.TextMatrix(i, .ColIndex("RemainingValue")))
If NetValu > RemainValu Then
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
Private Sub Option2_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

End Sub

Private Sub Option3_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If
If Option3.value = True Then
    XPTxtVal.Enabled = True
    Frame18.Visible = False

End If
CalCulteVAT 1
End Sub

Private Sub Option4_Click()
    Dim My_SQL As String
   
    currentname = DBCboClientName.text
    
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select Account_Code1,Emp_Name from TblEmployee where   (Account_Code1 <> N'""' AND NOT (Account_Code1 IS NULL)) "
Else
My_SQL = "  select Account_Code1,Emp_Namee from TblEmployee where   (Account_Code1 <> N'""' AND NOT (Account_Code1 IS NULL)) "
End If
My_SQL = My_SQL & " AND BranchId in(" & Current_branchSql & ")"
    fill_combo Me.DBCboClientName, My_SQL

    If txt_general_des.text = "" And Me.TxtModFlg <> "R" Then
        txt_general_des.text = Option4.Caption
    End If

    Fra(2).Visible = False
       lbl(47).Visible = False
        TxtAdvance.Visible = False
         DBCboClientName.text = currentname
End Sub

Private Sub Option5_Click()
    Dim My_SQL As String
    currentname = DBCboClientName.text
    If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select Account_Code,Emp_Name from TblEmployee    where (Account_Code <> N'""' AND NOT (Account_Code IS NULL)) "
    Else
    My_SQL = "  select Account_Code,Emp_Namee from TblEmployee    where (Account_Code <> N'""' AND NOT (Account_Code IS NULL)) "
    End If
    My_SQL = My_SQL & " AND BranchId in(" & Current_branchSql & ")"
    fill_combo Me.DBCboClientName, My_SQL

    If txt_general_des.text = "" And Me.TxtModFlg <> "R" Then
        txt_general_des.text = Option5.Caption
    End If

    Fra(2).Visible = True
       lbl(47).Visible = True
        TxtAdvance.Visible = True
           DBCboClientName.text = currentname
End Sub



Private Sub Option6_Click()
    Dim My_SQL As String
        currentname = DBCboClientName.text
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select Account_Code2,Emp_Name from TblEmployee where   (Account_Code2 <> N'""' AND NOT (Account_Code2 IS NULL)) "
 Else
 My_SQL = "  select Account_Code2,Emp_Namee from TblEmployee where   (Account_Code2 <> N'""' AND NOT (Account_Code2 IS NULL)) "
 End If
 My_SQL = My_SQL & " AND BranchId in(" & Current_branchSql & ")"
    fill_combo Me.DBCboClientName, My_SQL

    If txt_general_des.text = "" And Me.TxtModFlg <> "R" Then
        txt_general_des.text = Option6.Caption
    End If

    Fra(2).Visible = False
    lbl(47).Visible = False
        TxtAdvance.Visible = False
       DBCboClientName.text = currentname
End Sub

Private Sub Option7_Click()
    Dim My_SQL As String
        currentname = DBCboClientName.text
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select Account_Code3,Emp_Name from TblEmployee  where (Account_Code3 <> N'""' AND NOT (Account_Code3 IS NULL)) "
Else
My_SQL = "  select Account_Code3,Emp_Namee from TblEmployee  where (Account_Code3 <> N'""' AND NOT (Account_Code3 IS NULL)) "
End If
My_SQL = My_SQL & " AND BranchId in(" & Current_branchSql & ")"
    fill_combo Me.DBCboClientName, My_SQL

    If txt_general_des.text = "" And Me.TxtModFlg <> "R" Then
        txt_general_des.text = Option7.Caption
    End If

    Fra(2).Visible = False
    lbl(47).Visible = False
        TxtAdvance.Visible = False
            DBCboClientName.text = currentname
End Sub

Private Sub Option8_Click(Index As Integer)
CalCulteVAT
End Sub

Private Sub subContOpt_Click(Index As Integer)
DBCboClientName_Change
End Sub



Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Order_no_search.show
        If DCboCashType.ListIndex = 13 Or DCboCashType.ListIndex = 14 Then
            Order_no_search.RetrunType = 70
        Else
            Order_no_search.RetrunType = 2
        End If
    End If

End Sub

Public Sub TXT_order_no_Validate(Cancel As Boolean)

Dim s As String
Dim rsDummy  As ADODB.Recordset
If DCboCashType.ListIndex = 13 Or DCboCashType.ListIndex = 14 Then
    s = "Select * from TblLC Where TblLCID = " & val(txtTradingContractID)
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    Dim Dcombos As ClsDataCombos
     Set Dcombos = New ClsDataCombos
     Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    If Not rsDummy.EOF Then
        DcbAccount.BoundText = Trim(rsDummy!AcceptAccount_Code & "")
        'DBCboClientName.BoundText = 5
       ' DBCboClientName_Change
        DBCboClientName.BoundText = val(rsDummy!vendorID & "")
    End If
End If
End Sub

Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.text)
End Sub

Private Sub TxtAddOther_Change()
CalculteTaoals
End Sub

Private Sub TxtAddOther_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(TxtAddOther.text) + val(TxtPrePayd(4).text), 2) > Round(val(TxtAddOther2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’·Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
TxtAddOther.text = 0
TxtAddOther_Change
TxtAddOther.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtAdvance_Change()
If Me.TxtAdvance.text <> "" Then
If Me.TxtModFlg.text <> "R" Then
 EmpAdvanceRequest val(TxtAdvance.text)
End If
End If
End Sub

Private Sub TxtAdvance_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
  General_Search.send_form = "advreqPayment"
              Load General_Search
            General_Search.show
            
         'Load FrmEmpAdvanceSearch1
         '   FrmEmpAdvanceSearch1.Show vbModal
         '   FrmEmpAdvanceSearch1.returntype = 2
End If
End Sub

Private Sub txtAdvance1_Change()
'If Me.TxtModFlg.Text <> "R" Then
'TxtTotalsalary.Text = val(txtSalary.Text) + val(TxtSalEntitOther.Text) - val(TxtOther.Text) - val(txtAdvance1.Text) - val(TxtInsuranceValue.Text) + val(txttotal.Text)
'XPTxtVal.Text = val(TxtTotalsalary) + val(txtSalaryVocation.Text) + val(Me.txtValueTickt.Text)
'End If
Calculte
End Sub

Private Sub txtAdvance1_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(txtAdvance13.text) + val(txtAdvance1.text), 2) > Round(val(txtAdvance12.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’»Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
txtAdvance1.text = 0
txtAdvance1_Change
txtAdvance1.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TXTAdvanceTotal_Change()
CalculteTaoals
End Sub

Private Sub TXTAdvanceTotal_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(TXTAdvanceTotal.text) + val(TxtPrePayd(7).text), 2) > Round(val(TXTAdvanceTotal2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’·Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
TXTAdvanceTotal.text = 0
TXTAdvanceTotal_Change
TXTAdvanceTotal.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtCash_Change()
CalculteTaoals
End Sub

Private Sub TxtCash_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(TxtCash.text) + val(TxtPrePayd(9).text), 2) > Round(val(TxtCash2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’·Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
TxtCash.text = 0
TxtCash_Change
TxtCash.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtCurrencyRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
           LostAllFocus
        End If
End Sub
Sub CalCuteCurrency()
'Exit Sub
If val(TxtCurrencyRate.text) = 0 Then
TxtCurrencyRate.text = ""
End If
If val(val(TxtCurrencyRate.text)) <> 0 Then

'Text2.Text = Format(Text2.Text, "###.00")
XPTxtValE.text = Round(val(Format(XPTxtVal.text, "###.00")) / val(TxtCurrencyRate.text), 2)
XPTxtValE.text = Format(XPTxtValE.text, "#,##0.00")
Else
XPTxtValE.text = (XPTxtVal.text)
End If
If val(XPTxtValE.text) = 0 Then XPTxtVal.text = "": Exit Sub
End Sub
Private Sub TxtCurrencyRate_KeyUp(KeyCode As Integer, Shift As Integer)
CalCuteCurrency
End Sub

Private Sub TxtCustCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer
    Dim ID As Double
Dim Account_code As String
Dim lblflag As Integer
    If KeyAscii = vbKeyReturn Then
        
If val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 1 Then
        
        GetCustomersDetail CUSTID, , TxtCustCode.text, DCboCashType.ListIndex + 1
        DBCboClientName.BoundText = CUSTID

ElseIf val(DCboCashType.ListIndex) = 4 Then


        If Option4.value = True Then
        lblflag = 1
       ElseIf Option5.value = True Then
        lblflag = 0
       
       ElseIf Option6.value = True Then
        lblflag = 2
      ElseIf Option7.value = True Then
        lblflag = 3
       End If


Me.DcbEmpBranch.BoundText = GetEmployeeBranch(EmpIDD)

  GetEmployeeIDFromCode TxtCustCode.text, , , , , lblflag, Account_code
        DBCboClientName.BoundText = Account_code
 ElseIf val(Me.DCboCashType.ListIndex) = 5 Then
DBCboClientName.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtCustCode.text)
ElseIf val(Me.DCboCashType.ListIndex) = 3 Then
    If KeyAscii = vbKeyReturn Then
    If TxtCustCode.text <> "" Then
GetCodeIDProject ID, TxtCustCode.text
DBCboClientName.BoundText = ID
    End If
    End If
End If

    End If

End Sub
Function CheckVacationPayed(Optional ID As Double, Optional ByRef salary As Double, Optional ByRef SalEntitOther As Double, Optional ByRef SalaryVocation As Double, Optional ByRef ValueTickt As Double _
, Optional ByRef other As Double, Optional ByRef Advance As Double, Optional ByRef InsuranceValue As Double) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim StrSQL As String
Dim sql As String
Dim Salarall As Double
Dim ch8 As Integer
StrSQL = "select sum(InsuranceValue) as SmInsuranceValue, Sum(Advance1) as SmAdvance1, Sum(Other)as SmOther ,sum(Salary33) as SmSalary ,sum(SalEntitOther)as SmSalEntitOther,sum(SalaryVocation)as SmSalaryVocation ,Sum(ValueTickt)as SmValueTickValueTickt From Notes where NoteType=5 and ( cashingtype<=10 ) and Due=" & val(TxtDue.text) & " "
If Me.TxtModFlg.text = "E" Then
StrSQL = StrSQL & " and NoteID <>" & val(XPTxtID.text) & ""
End If
Rs8.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
salary = IIf(IsNull(Rs8("SmSalary").value), 0, Rs8("SmSalary").value)
SalEntitOther = IIf(IsNull(Rs8("SmSalEntitOther").value), 0, Rs8("SmSalEntitOther").value)
SalaryVocation = IIf(IsNull(Rs8("SmSalaryVocation").value), 0, Rs8("SmSalaryVocation").value)
ValueTickt = IIf(IsNull(Rs8("SmValueTickValueTickt").value), 0, Rs8("SmValueTickValueTickt").value)
Dim s As String
Dim rsDummy As New ADODB.Recordset

s = "Select IsNull(Booked,0) as Booked,ValueTickt,PaymentRecommended From TblVocationEntitlements Where Id = " & val(TxtDue.text) & " and not (NoteSerial is null)"
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly, adCmdText
If Not rsDummy.EOF Then
    If Not rsDummy!Booked Then
      '  ValueTickt = val(rsDummy!ValueTickt & "")
    End If
    txtPaymentRecommended = rsDummy!PaymentRecommended & ""
End If

other = IIf(IsNull(Rs8("SmOther").value), 0, Rs8("SmOther").value)
Advance = IIf(IsNull(Rs8("SmAdvance1").value), 0, Rs8("SmAdvance1").value)
InsuranceValue = IIf(IsNull(Rs8("SmInsuranceValue").value), 0, Rs8("SmInsuranceValue").value)
Set Rs3 = New ADODB.Recordset
sql = "Select * from tblVocationEntitlements where ID=" & ID & " and not (NoteSerial is null)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs3.RecordCount > 0 Then
 ch8 = IIf(IsNull(Rs3("ch8").value), 0, Rs3("ch8").value)
 If ch8 = 1 Then
Salarall = IIf(IsNull(Rs3("Salary").value), 0, Rs3("Salary").value) + IIf(IsNull(Rs3("PreSalary").value), 0, Rs3("PreSalary").value)
Else
Salarall = IIf(IsNull(Rs3("Salary").value), 0, Rs3("Salary").value)
End If
If Salarall <> salary Then
CheckVacationPayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("SalEntitOther").value), 0, Rs3("SalEntitOther").value) <> SalEntitOther Then
CheckVacationPayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("SalaryVocation").value), 0, Rs3("SalaryVocation").value) <> SalaryVocation Then
CheckVacationPayed = False
Exit Function
ElseIf IIf(IsNull(Rs8("SmValueTickValueTickt").value), 0, Rs8("SmValueTickValueTickt").value) <> ValueTickt Then
CheckVacationPayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("Other").value), 0, Rs3("Other").value) <> other Then
CheckVacationPayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("Advance").value), 0, Rs3("Advance").value) <> Advance Then
CheckVacationPayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("InsuranceValue").value), 0, Rs3("InsuranceValue").value) <> InsuranceValue Then
CheckVacationPayed = False
ElseIf IIf(IsNull(Rs3("ValueTickt").value), 0, Rs3("ValueTickt").value) <> ValueTickt Then
CheckVacationPayed = False

Exit Function
Else
CheckVacationPayed = True
Exit Function
End If
CheckVacationPayed = False
Exit Function
End If
Else
CheckVacationPayed = False
Exit Function
End If
End Function
Function CheckEndSerivecePayed(Optional ByRef total As Double, Optional ByRef Sal As Double, Optional ByRef Custom2 As Double, Optional ByRef TicketValue As Double _
, Optional ByRef CusTiket As Double, Optional ByRef AddOther As Double, Optional ByRef ValEndService As Double _
, Optional ByRef AdvanceTotal As Double, Optional ByRef VlueVaction As Double, Optional ByRef Cash As Double, Optional ByRef Discounts As Double, Optional ByRef DisSalary As Double) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim StrSQL As String
Dim sql As String
StrSQL = "select sum(total) as Smtotal, Sum(Sal) as SmSal, Sum(Custom2)as SmCustom2 ,sum(TicketValue) as SmTicketValue ,sum(CusTiket)as SmCusTiket,sum(AddOther)as SmAddOther ,Sum(ValEndService)as SmValEndService"
StrSQL = StrSQL & ",sum(AdvanceTotal)as SmAdvanceTotal,sum(VlueVaction)as SmVlueVaction,sum(Cash)as SmCash ,sum(Discounts)as SmDiscounts ,sum(DisSalary) as SmDisSalary From Notes where NoteType=5 and ( cashingtype=10 ) and TxtEndService=" & val(TxtEndService.text) & " "
If Me.TxtModFlg.text = "E" Then
StrSQL = StrSQL & " and NoteID <>" & val(XPTxtID.text) & ""
End If
Rs8.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
total = IIf(IsNull(Rs8("Smtotal").value), 0, Rs8("Smtotal").value)
Sal = IIf(IsNull(Rs8("SmSal").value), 0, Rs8("SmSal").value)
Custom2 = IIf(IsNull(Rs8("SmCustom2").value), 0, Rs8("SmCustom2").value)
TicketValue = IIf(IsNull(Rs8("SmTicketValue").value), 0, Rs8("SmTicketValue").value)
CusTiket = IIf(IsNull(Rs8("SmCustom2").value), 0, Rs8("SmCustom2").value)
AddOther = IIf(IsNull(Rs8("SmAddOther").value), 0, Rs8("SmAddOther").value)
ValEndService = IIf(IsNull(Rs8("SmValEndService").value), 0, Rs8("SmValEndService").value)
AdvanceTotal = IIf(IsNull(Rs8("SmAdvanceTotal").value), 0, Rs8("SmAdvanceTotal").value)
VlueVaction = IIf(IsNull(Rs8("SmVlueVaction").value), 0, Rs8("SmVlueVaction").value)
Cash = IIf(IsNull(Rs8("SmCash").value), 0, Rs8("SmCash").value)
Discounts = IIf(IsNull(Rs8("SmDiscounts").value), 0, Rs8("SmDiscounts").value)
DisSalary = IIf(IsNull(Rs8("SmDisSalary").value), 0, Rs8("SmDisSalary").value)

Set Rs3 = New ADODB.Recordset
sql = "Select * from End_of_service where ID=" & val(TxtEndService.text) & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs3.RecordCount > 0 Then

If IIf(IsNull(Rs3("NetEnd").value), IIf(IsNull(Rs3("total").value), 0, Rs3("total").value), Rs3("NetEnd").value) <> total Then
CheckEndSerivecePayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("LastMonth").value), 0, Rs3("LastMonth").value) <> Sal Then
CheckEndSerivecePayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("Custom").value), 0, Rs3("Custom").value) <> Custom2 Then
CheckEndSerivecePayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("Ticket").value), 0, Rs3("Ticket").value) <> TicketValue Then
CheckEndSerivecePayed = False
Exit Function
'ElseIf IIf(IsNull(Rs3("CusTiket").value), 0, Rs3("CusTiket").value) <> CusTiket Then
'CheckEndSerivecePayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("AddOther").value), 0, Rs3("AddOther").value) <> AddOther Then
CheckEndSerivecePayed = False
Exit Function
'ElseIf IIf(IsNull(Rs3("EndService").value), 0, Rs3("EndService").value) <> ValEndService Then
'CheckEndSerivecePayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("TotalAdvance").value), 0, Rs3("TotalAdvance").value) <> AdvanceTotal Then
CheckEndSerivecePayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("TxtVlueVaction").value), 0, Rs3("TxtVlueVaction").value) <> VlueVaction Then
CheckEndSerivecePayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("TotalCash").value), 0, Rs3("TotalCash").value) <> Cash Then
CheckEndSerivecePayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("Discounts").value), 0, Rs3("Discounts").value) <> Discounts Then
CheckEndSerivecePayed = False
Exit Function
ElseIf IIf(IsNull(Rs3("DisSalary").value), 0, Rs3("DisSalary").value) <> DisSalary Then
CheckEndSerivecePayed = False
Exit Function
Else
CheckEndSerivecePayed = True
Exit Function
End If
CheckEndSerivecePayed = False
Exit Function
End If
Else
CheckEndSerivecePayed = False
Exit Function
End If
End Function

Private Sub TxtCusTiket_Change()
CalculteTaoals
End Sub

Private Sub txtCustom2_Change()
CalculteTaoals
End Sub

Private Sub txtCustom2_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(txtCustom2.text) + val(TxtPrePayd(2).text), 2) > Round(val(txtCustom.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’·Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
txtCustom2.text = 0
txtCustom2_Change
txtCustom2.SetFocus
Exit Sub
End If
End If
End Sub



Sub DuVac()
Dim Salar As Double
Dim SalEntitOther As Double
Dim SalaryVocation As Double
Dim ValueTickt As Double
Dim other As Double
Dim Advance As Double
Dim InsuranceValue As Double
'Dim ValueTickt As Double
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
If DCboCashType.ListIndex = 8 Then
 
 txtValueTickt3.text = 0
 txtValueTickt2.text = 0
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
 If CheckVacationPayed(val(TxtDue.text), Salar, SalEntitOther, SalaryVocation, ValueTickt, other, Advance, InsuranceValue) = True Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox " „ ”œ«œ Â–Â «·Õ—ﬂ… „‰ ﬁ»· "
 Else
 MsgBox "It Was Paid"
 End If
 txtSalary3.text = 0
 TxtSalEntitOther3.text = 0
 txtSalaryVocation3.text = 0
 txtValueTickt3.text = 0
 TxtOther3.text = 0
 txtAdvance13.text = 0
 TxtInsuranceValue3.text = 0
 
  DcboEmpName.BoundText = 0
 Me.dcBranch1.BoundText = 0
 TxtSearchCode.text = "'"
 TxtSalary.text = 0
 TxtSalEntitOther.text = 0
 Txtother.text = 0
 txtAdvance1.text = 0
Me.txtValueTickt.text = 0
 txtSalaryVocation.text = 0
 Me.TxtInsuranceValue.text = 0
 TxtTotalsalary = 0
XPTxtVal.text = 0
 txtSalary2.text = 0
 TxtSalEntitOther2.text = 0
 TxtOther2.text = 0
 txtAdvance12.text = 0
Me.txtValueTickt2.text = 0
 txtSalaryVocation2.text = 0
 Me.TxtInsuranceValue2.text = 0
 TxtTotalsalary2 = 0


 Exit Sub
  Else
  txtSalary3.text = Salar
 TxtSalEntitOther3.text = SalEntitOther
 txtSalaryVocation3.text = SalaryVocation
 txtValueTickt3.text = ValueTickt
 TxtOther3.text = other
 txtAdvance13.text = Advance
 TxtInsuranceValue3.text = InsuranceValue
 End If

 End If
Dim salary  As Double
Dim PreSalary As Double
'Dim Other  As Double
'Dim Advance  As Double
Dim EmpID As Integer
' Dim ValueTickt  As Double
'Dim SalaryVocation As Double
'Dim InsuranceValue As Double
Dim BranchID As Integer

If 1 = 1 Then
  GetVocationEntitlements val(TxtDue), BranchID, EmpID, salary, SalEntitOther, other, Advance, ValueTickt, SalaryVocation, InsuranceValue, PreSalary
 
Dim s As String
Dim rsDummy As New ADODB.Recordset

s = "Select IsNull(Booked,0) as Booked,ValueTickt From TblVocationEntitlements Where Id = " & val(TxtDue.text) & " and not (NoteSerial is null)"
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly, adCmdText
If Not rsDummy.EOF Then
    If Not rsDummy!Booked Then
        ValueTickt = val(rsDummy!ValueTickt & "")
    Else
    ValueTickt = 0
    End If
End If
 DcboEmpName.BoundText = EmpID
 Me.dcBranch1.BoundText = BranchID
 salary = salary + PreSalary
 TxtSalary.text = Round(salary - val(txtSalary3.text), 2)
 TxtSalEntitOther.text = Round(SalEntitOther - val(TxtSalEntitOther3.text), 2)
 Txtother.text = Round(other - val(TxtOther3.text), 2)
 txtAdvance1.text = Round(Advance - val(txtAdvance13.text), 2)
Me.txtValueTickt.text = Round(ValueTickt - val(Me.txtValueTickt3.text), 2)
 txtSalaryVocation.text = Round(SalaryVocation - val(txtSalaryVocation3.text), 2)
 
 Me.TxtInsuranceValue.text = Round(InsuranceValue - val(Me.TxtInsuranceValue3.text), 2)
  txtSalary2.text = salary
 TxtSalEntitOther2.text = SalEntitOther
 TxtOther2.text = other
 txtAdvance12.text = Advance
Me.txtValueTickt2.text = ValueTickt
 txtSalaryVocation2.text = SalaryVocation
 Me.TxtInsuranceValue2.text = InsuranceValue
TxtTotalsalary2.text = val(salary) + val(SalEntitOther) - val(other) - val(Advance) - val(InsuranceValue)
 TxtTotalsalary.text = val(TxtSalary.text) + val(TxtSalEntitOther.text) - val(Txtother.text) - val(txtAdvance1.text) - val(TxtInsuranceValue.text)
 XPTxtVal.text = 0
XPTxtVal.text = val(TxtTotalsalary.text) + val(txtSalaryVocation.text) + val(Me.txtValueTickt.text)
XPTxtVal.text = Round(val(XPTxtVal.text), 2)
End If
End If
End If
End Sub

Private Sub TxtDue_Change()
If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
DuVac
End If
End Sub

Sub EndServicess()
If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
 Dim EmpID  As Double
 Dim Discounts As Double
 Dim total  As Double
 Dim EndService As Double
 Dim LastMonth  As Double
 Dim Ticket  As Double
 Dim Custom As Double
 Dim net  As Double
 Dim CusTiket As Double
 Dim TotalAdvance As Double
 Dim TxtVlueVaction As Double
 Dim TotalCash As Double
 Dim LastTotal As Double
 Dim BranchID As Integer
 Dim AddOther As Double
 Dim DiffTekit As Double
 Dim TicktConract As Double
 Dim DisSalary As Double
 'Dim Total As Double
 Dim Sal As Double
 Dim Custom2 As Double
 Dim TicketValue As Double
' Dim CusTiket As Double
' Dim AddOther As Double
 Dim ValEndService As Double
 Dim AdvanceTotal As Double
 Dim VlueVaction As Double
 'Dim VlueVaction As Double
 Dim Cash As Double
 'Dim DisSalary As Double
' Dim Discounts As Double
If DCboCashType.ListIndex = 10 Then
If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
'If ChekPayment() = True Then
If CheckEndSerivecePayed(total, Sal, Custom2, TicketValue, CusTiket, AddOther, ValEndService, AdvanceTotal, VlueVaction, Cash, Discounts, DisSalary) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Â–Â «·⁄„·Ì…  „ ”œ«œÂ« „”»ﬁ« Ì—ÃÏ «Œ Ì«— ⁄„·Ì… «Œ—Ï"
Else
MsgBox "This process is already paid"
End If
TxtTicktConract.text = 0
DcbEmpEndService.BoundText = 0
  Me.DcbBranchEndServ.BoundText = 0
  txtTotal.text = 0
  Me.txtSal.text = 0
  Me.TxtVlueVaction.text = 0
  txtTicketValue.text = 0
  txtCustom.text = 0
  txtNet.text = 0
  TXTAdvanceTotal.text = 0
  TxtCash.text = 0
  TXTLastTotal.text = 0
  XPTxtVal.text = 0
    txtTotal2.text = 0
  Me.txtSal2.text = 0
  Me.TxtVlueVaction2.text = 0
  txtTicketValue2.text = 0
  txtCustom2.text = 0
  txtNet2.text = 0
  TXTAdvanceTotal2.text = 0
  TxtCash2.text = 0
  TXTLastTotal2.text = 0
  TxtValEndService2.text = 0
  TxtValEndService.text = 0
  TxtCusTiket.text = 0
  TxtCusTiket2.text = 0
  TxtAddOther2.text = 0
  TxtAddOther.text = 0
  TxtPrePayd(12).text = 0
  TxtPrePayd(13).text = 0
  TxtPrePayd(14).text = 0
  TxtPrePayd(15).text = 0
  TxtPrePayd(16).text = 0
  
  TxtPrePayd(0).text = 0
  TxtPrePayd(1).text = 0
  TxtPrePayd(2).text = 0
  TxtPrePayd(3).text = 0
  TxtPrePayd(4).text = 0
  TxtPrePayd(5).text = 0
  TxtPrePayd(6).text = 0
  TxtPrePayd(7).text = 0
  TxtPrePayd(8).text = 0
  TxtPrePayd(9).text = 0
  TxtPrePayd(10).text = 0
  TxtPrePayd(11).text = 0
  TxtTotalDis2.text = 0
Exit Sub
Else
  TxtPrePayd(0).text = total
  TxtPrePayd(1).text = Sal
  TxtPrePayd(2).text = Custom2
  TxtPrePayd(3).text = TicketValue
  TxtPrePayd(4).text = AddOther
  TxtPrePayd(5).text = total + Sal + Custom2 + TicketValue + AddOther
  TxtPrePayd(6).text = ValEndService
  TxtPrePayd(7).text = AdvanceTotal
  TxtPrePayd(8).text = VlueVaction
  TxtPrePayd(9).text = Cash
  TxtPrePayd(10).text = Discounts
  TxtPrePayd(15).text = DisSalary
  TxtPrePayd(11).text = ValEndService + AdvanceTotal + VlueVaction + Cash + Discounts + DisSalary
  TxtTotlPaidEndSer.text = val(TxtPrePayd(5).text) - val(TxtPrePayd(11).text)
End If
End If


  GetEnd_Service val(TxtEndService.text), BranchID, EmpID, total, LastMonth, Ticket, Custom, net, TotalAdvance, TxtVlueVaction, TotalCash, LastTotal, EndService, CusTiket, AddOther, DiffTekit, Discounts, TicktConract, DisSalary
  DcbEmpEndService.BoundText = EmpID
  Me.DcbBranchEndServ.BoundText = BranchID
  TxtValEndService2.text = 0 ' EndService
  TxtValEndService.text = 0 ' EndService
  TxtTicktConract.text = TicktConract
  TxtPrePayd(16).text = DisSalary - val(TxtPrePayd(15).text)
  TxtPrePayd(14).text = TxtPrePayd(16).text
  
  TxtPrePayd(13).text = Discounts - val(TxtPrePayd(10).text)
  TxtPrePayd(12).text = TxtPrePayd(13).text
  txtTotal.text = total - val(TxtPrePayd(0).text)
  txtTotal2.text = total
  Me.txtSal.text = LastMonth - val(TxtPrePayd(1).text)
  Me.txtSal2.text = LastMonth
  txtTicketValue.text = Ticket - val(TxtPrePayd(3).text)
  txtTicketValue2.text = Ticket
  txtCustom.text = Custom
  txtCustom2.text = Custom - val(TxtPrePayd(2).text)
  TxtCusTiket2.text = CusTiket
 ' TxtCusTiket.Text = CusTiket + DiffTekit
  TxtAddOther2.text = AddOther - val(TxtPrePayd(4).text)
  TxtAddOther.text = AddOther
  txtNet2.text = val(TxtAddOther2.text) + val(TxtCusTiket2.text) + val(txtCustom2.text) + val(txtTicketValue2.text) + val(txtSal2.text) + val(txtTotal2.text)
  txtNet.text = val(TxtAddOther.text) + val(TxtCusTiket.text) + val(txtCustom2.text) + val(txtTicketValue.text) + val(txtSal.text) + val(txtTotal2.text)
  TXTAdvanceTotal.text = TotalAdvance - val(TxtPrePayd(7).text)
  TXTAdvanceTotal2.text = TotalAdvance
  Me.TxtVlueVaction.text = TxtVlueVaction - val(TxtPrePayd(8).text)
  Me.TxtVlueVaction2.text = TxtVlueVaction
  TxtCash.text = TotalCash - val(TxtPrePayd(9).text)
  TxtCash2.text = TotalCash
 ' TxtTotalDis2.Text = val(TXTAdvanceTotal2.Text) + val(Me.TxtVlueVaction2.Text) + val(TxtCash2.Text) + val(TxtDiscounts2.Text)
 ' TxtTotalDis.Text = val(TXTAdvanceTotal.Text) + val(Me.TxtVlueVaction.Text) + val(TxtCash.Text) + val(TxtDiscounts.Text)
 ' TXTLastTotal2.Text = val(txtnet2.Text) + val(TxtValEndService2.Text) - val(TxtTotalDis2.Text)
 '  TXTLastTotal.Text = val(txtnet.Text) + val(TxtValEndService.Text) - val(TxtTotalDis.Text)
  'XPTxtVal.Text = val(TxtTotal.Text) + val(TxtTicketValue.Text) + val(TxtCash.Text) + val(TxtAddOther.Text) + val(txtSal.Text)
  'XPTxtVal.Text = val(XPTxtVal.Text) - val(TXTDiscounts.Text) - val(TxtCash.Text) - val(TXTAdvanceTotal.Text)
  CalculteTaoals
   
End If
End If
End Sub

Private Sub TxtDue_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
     FrmSearchVocationEntitlement.Index = 2
            Load FrmSearchVocationEntitlement
          FrmSearchVocationEntitlement.show vbModal
End If

End Sub

Private Sub TxtDue_Validate(Cancel As Boolean)
DuVac
End Sub

Private Sub TxtEndService_Change()
If DCboCashType.ListIndex = 12 Then

Else
EndServicess
End If
End Sub
Sub GetVAT(Optional ID As Double)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "select * from TblVATAvowal where ID=" & ID & " and (Payed is null or Payed=0)"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
XPTxtVal.text = IIf(IsNull(rs2("TxtNetVat").value), 0, rs2("TxtNetVat").value) ''- IIf(IsNull(Rs2("VATPurchasesTotal").value), 0, Rs2("VATPurchasesTotal").value)
If SystemOptions.UserInterface = ArabicInterface Then
txt_general_des.text = "”œ«œ «·ﬁÌ„… «·„÷«›… ⁄‰ «·› —Â „‰ "
txt_general_des.text = txt_general_des.text & " " & IIf(IsNull(rs2("DateFrom").value), "", rs2("DateFrom").value)
txt_general_des.text = txt_general_des.text & " " & " «·Ï" & IIf(IsNull(rs2("DateFrom").value), "", rs2("DateFrom").value)
Else
txt_general_des.text = "VAT of Period from  "
txt_general_des.text = txt_general_des.text & " " & IIf(IsNull(rs2("DateFrom").value), "", rs2("DateFrom").value)
txt_general_des.text = txt_general_des.text & " " & " to" & " " & IIf(IsNull(rs2("DateFrom").value), "", rs2("DateFrom").value)
End If
DcbCurrency_Change
Else
XPTxtVal.text = 0
txt_general_des.text = ""
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " «ﬂœ „‰ —ﬁ„ «·«ﬁ—«— «Ê Â–Â «·Õ—ﬂ…  „ œ›⁄Â« „”»ﬁ"
Else
MsgBox "Check the process number or this is process already payed"
End If
End If
End Sub
Function ChekPayment() As Boolean
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
Dim sql As String
ChekPayment = False
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "Select id from  End_of_service where id=" & val(TxtEndService.text) & " and PaymPaid=1 "
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekPayment = True
Else
ChekPayment = False
End If
End If
End Function

Private Sub TxtEndService_KeyPress(KeyAscii As Integer)
If DCboCashType.ListIndex = 12 Then
If KeyAscii = vbKeyReturn Then
If val(val(TxtEndService.text)) <> 0 Then
GetVAT val(TxtEndService.text)
End If
End If
End If
End Sub

Private Sub TxtEndService_KeyUp(KeyCode As Integer, Shift As Integer)
If DCboCashType.ListIndex = 12 Then
Else
If KeyCode = vbKeyF3 Then

   Unload FrmEnserviceSearch
    Load FrmEnserviceSearch
    FrmEnserviceSearch.Index = 2
    FrmEnserviceSearch.show

End If
End If
End Sub

Private Sub TxtInsuranceValue_Change()
'If Me.TxtModFlg.Text <> "R" Then
'TxtTotalsalary.Text = val(txtSalary.Text) + val(TxtSalEntitOther.Text) - val(TxtOther.Text) - val(txtAdvance1.Text) - val(TxtInsuranceValue.Text)
'XPTxtVal.Text = val(TxtTotalsalary) + val(txtSalaryVocation.Text) + val(Me.txtValueTickt.Text)
'End If
Calculte
End Sub

Private Sub TxtInsuranceValue_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(TxtInsuranceValue3.text) + val(TxtInsuranceValue.text), 2) > Round(val(TxtInsuranceValue2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’»Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
TxtInsuranceValue.text = 0
TxtInsuranceValue_Change
TxtInsuranceValue.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
'ReloadContracR
    Select Case Me.TxtModFlg.text

        Case "R"
        Command6.Enabled = False
        GRID1.Enabled = True
        Check18.Enabled = False
Command10(0).Enabled = False
Command10(1).Enabled = False
        VSFlexGrid1.Enabled = True
        Check1.Enabled = False
        
            Frame2.Enabled = False

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Payments"
            Else
                Me.Caption = "«·„œ›Ê⁄« "
     
            End If
            Command7.Enabled = False
            Command3.Enabled = False
            Command6.Enabled = False

            Command2.Enabled = False
            Command3.Enabled = False
dcBranch.Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            XPTxtVal.locked = True
            XPDtbTrans.Enabled = False
            XPMTxtRemarks.locked = True
            DBCboClientName.locked = True
            DcboBox.locked = True
            DCboCashType.locked = True
            Me.CboPayMentType.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

            Fra(0).Enabled = False
            ChkTrans.Enabled = False

        Case "N"
        Command7.Enabled = False
        Command3.Enabled = False
        Command6.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
 Command6.Enabled = False
        GRID1.Enabled = True
        Check18.Enabled = True
        
        Command10(0).Enabled = False
        Command10(1).Enabled = False
        VSFlexGrid1.Enabled = True
        Check1.Enabled = True

dcBranch.Enabled = True
            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Payments (New)"
            Else
                Me.Caption = "«·„œ›Ê⁄« ( ÃœÌœ )"
        
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
            XPDtbTrans.Enabled = True
            XPTxtVal.locked = False
            XPMTxtRemarks.locked = False
            DBCboClientName.locked = False
            DcboBox.locked = False
            Me.CboPayMentType.locked = False
            XPDtbTrans.value = Date
            DCboCashType.locked = False
            DBCboClientName.locked = False
        If called = False Then
            DCboCashType.ListIndex = 1
          End If
            Fra(0).Enabled = True
            ChkTrans.Enabled = True

        Case "E"
     '   If val(DCboCashType.ListIndex) = 9 Or val(DCboCashType.ListIndex) = 8 Or val(DCboCashType.ListIndex) = 7 Then
     '   If FrmEmpSalary6.ClearPayment = True Then
     '   Cmd(2).Enabled = True
     '   Cmd(3).Enabled = False
     '   Else
     '   Cmd(3).Enabled = True
     '   Cmd(2).Enabled = False
     '   End If
     '   Else
     '   Cmd(3).Enabled = True
     '   Me.Cmd(2).Enabled = True
     '   End If
      Cmd(3).Enabled = True
      Me.Cmd(2).Enabled = True
        Command2.Enabled = True
           Command7.Enabled = True
            Command3.Enabled = True
            Command6.Enabled = True
        Command3.Enabled = True
dcBranch.Enabled = True
            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Payments (Edit)"
            Else
                Me.Caption = "«·„œ›Ê⁄« (  ⁄œÌ· )"
        
            End If
            Check1.Enabled = False
            Command10(0).Enabled = False
            Command10(1).Enabled = False
            FlgBillBuy = False
            FlgBill = False
            FlgBillProject = False
  Command6.Enabled = True
        GRID1.Enabled = False
        Check18.Enabled = False
        VSFlexGrid1.Enabled = False
          Command10(0).Enabled = True
   Command10(1).Enabled = True
        Check1.Enabled = False
   
        
         '   Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            XPTxtVal.locked = False
            XPDtbTrans.Enabled = True
            DcboBox.locked = False
            XPMTxtRemarks.locked = False
            DBCboClientName.locked = False
            DCboCashType.locked = False
            DBCboClientName.locked = False
            Me.CboPayMentType.locked = False
            Fra(0).Enabled = True
            ChkTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub



Private Sub txtNet_Change()
CalculteTaoals
End Sub

Private Sub TxtOrder_Change()
TxtDue.text = ""
DuVac

If TxtOrder.text = "" Then
DBCboClientName.Enabled = True
Frame2.Enabled = True
DCboCashType.Enabled = True
End If

TxtEndService_Change
If TxtOrder.text <> "" Then
DBCboClientName.Enabled = False
Frame2.Enabled = False
DCboCashType.Enabled = False
Dim CusID As Double
Dim Account_code As String
Dim frmtype As Integer
Dim Type1 As Integer
Dim txtperson As String
Dim des As String
Dim EmpID As Integer
Dim NotValue As Double
Dim orderNo As String
Dim basedOn As Integer
Dim CurrcyID As Integer
Dim valuee As Double
Dim Rate As Double

Dim Transaction_ID As Long
If Me.TxtModFlg.text <> "R" And Me.TxtModFlg.text <> "" Then
TxtDue.text = ""
TxtEndService.text = ""
Dim salary_or_advance As Integer
  OrderExchange TxtOrder, Type1, txtperson, des, Price, EmpID, basedOn, orderNo, Transaction_ID, CusID, frmtype, Account_code, CurrcyID, Rate, valuee, salary_or_advance
  
    
NotValue = GetVal((Me.TxtOrder.text), val(Me.XPTxtID), 5)
If Price = -1 Then Exit Sub

 

If NotValue < Price Then
DCboCashType.ListIndex = frmtype
DCboCashType_Change

If frmtype = 5 Then
DBCboClientName.BoundText = Account_code
Else
DBCboClientName.BoundText = CusID
End If


   If frmtype = 4 Then

        If IsNull(salary_or_advance) Then
            Option4.value = False: Option5.value = False
        ElseIf (salary_or_advance) = 0 Then
            Option4.value = True
            Option4_Click
        ElseIf (salary_or_advance) = 1 Then
            Option5.value = True
            Option5_Click
        ElseIf (salary_or_advance) = 2 Then
            Option6.value = True
            Option6_Click
        ElseIf (salary_or_advance) = 3 Then
            Option7.value = True
            Option7_Click
        End If
        DBCboClientName.text = txtperson
        
        End If
        
DBCboClientName_Change
Me.DcbCurrency.BoundText = CurrcyID
TxtCurrencyRate.text = Rate
XPTxtVal.text = Price - NotValue
XPTxtValE.text = val(XPTxtVal.text) / Rate
txt_general_des.text = des
XPMTxtRemarks.text = txtperson

CboPayMentType.ListIndex = Type1


If basedOn = 3 Then
Me.DCboCashType.ListIndex = 8
If Transaction_ID = 0 Then
TxtDue.text = orderNo

Else
TxtDue.text = Transaction_ID
DuVac
End If
ElseIf basedOn = 4 Then
Me.DCboCashType.ListIndex = 10
If Transaction_ID = 0 Then
TxtEndService.text = orderNo
Else
TxtEndService.text = Transaction_ID
End If

End If


Else
TxtOrder.text = ""
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·ﬁœ  „ ’—› ﬁÌ„… Â–« «·”‰œ »«·ﬂ«„· Ì—ÃÏ «Œ Ì«— ÿ·» ’—› «Œ—"
Else
MsgBox " Not Foud Value Please Select Another Order"
End If
End If
End If

End If
End Sub
Function ReriveAccountCode(Optional EmpID As Integer = 0) As String
 Dim My_SQL As String
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  
        My_SQL = "  select Account_Code from TblEmployee    where Emp_ID= " & EmpID & ""
        rs.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If rs.RecordCount > 0 Then
      
   ReriveAccountCode = IIf(IsNull(rs("Account_Code").value), "", (rs("Account_Code").value))
 
Else
ReriveAccountCode = ""
End If
End Function
Public Sub EmpAdvanceRequest(Serial1 As Integer)
      Dim i As Integer
                                
     Dim StrSQL  As String
     Dim rs As ADODB.Recordset
     Dim RsDetails As ADODB.Recordset
    Set rs = New ADODB.Recordset
  StrSQL = " SELECT     * "
StrSQL = StrSQL & " FROM     TblEmpAdvanceRequest where AdvanceID =" & Serial1 & ""

     If CheckAprroveScreen("FrmEmpsAdvanceRequest") = True Then

  StrSQL = StrSQL & "  and Approved = 1"
  End If
  
  
      rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    'rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (rs.RecordCount) > 0 Then
    Me.DBCboClientName.BoundText = ReriveAccountCode(val(IIf(IsNull(rs("Emp_ID").value), 0, (rs("Emp_ID").value))))
       txt_general_des.text = IIf(IsNull(rs("DiscountDES").value), "", rs("DiscountDES").value)
       
      TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), 0, (rs("PaymentCounts").value))
   Me.CboYear.text = IIf(IsNull(rs("FirstYearPayment").value), 0, (rs("FirstYearPayment").value)) ' rs("FirstYearPayment").value
       CmbMonth.ListIndex = val(IIf(IsNull(rs("FirstMonthPayment").value), 0, (rs("FirstMonthPayment").value))) - 1
       
       
       
       
         XPTxtVal.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
       Me.ChkSaleryDis.value = IIf(rs("AutoDiscount").value = True, vbChecked, vbUnchecked)
      
         Set RsDetails = New ADODB.Recordset
    StrSQL = "Select * From  TblEmpAdvanceRequestDetails Where AdvanceID =" & Serial1 & ""
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = FG.FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        Fra(2).Visible = True
        lbl(47).Visible = True
        TxtAdvance.Visible = True
        RsDetails.MoveFirst
        FG.rows = FG.FixedRows + RsDetails.RecordCount

        For i = Me.FG.FixedRows To FG.rows - 1
            FG.TextMatrix(i, FG.ColIndex("PartNO")) = IIf(IsNull(RsDetails("PartNO").value), "", (RsDetails("PartNO").value))
            FG.TextMatrix(i, FG.ColIndex("PartValue")) = IIf(IsNull(RsDetails("PartValue").value), "", (RsDetails("PartValue").value))
            FG.TextMatrix(i, FG.ColIndex("PartDate")) = IIf(IsNull(RsDetails("PartDate").value), Date, (RsDetails("PartDate").value))
            RsDetails.MoveNext
        Next i

    End If

    RsDetails.Close
    Set RsDetails = Nothing
    Set rs = Nothing
               
    End If
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
Cmd_Click (11)
End If
'    rs.Close

End Sub

Private Sub TxtOrder_KeyUp(KeyCode As Integer, Shift As Integer)
TxtDue.text = ""
If Me.TxtModFlg.text <> "R" Then
If KeyCode = vbKeyF3 Then
 Load FrmReqExchangeSearch
             FrmReqExchangeSearch.show
             FrmReqExchangeSearch.lbltype = 1
            
End If
End If
End Sub

Private Sub TxtOrderSuppler_Change()
Dim YraID As Integer
Dim MonthID As Integer
Dim BranchID As Integer
Dim AllID  As String
AllID = GetExchangReq(val(TxtOrderSuppler.text), YraID, MonthID, BranchID)
dcDur.BoundText = YraID
dcMontth.BoundText = MonthID
DcbBrReq.BoundText = BranchID
End Sub

Private Sub TxtOrderSuppler_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
 Unload FrmSearch_Request
                       FrmSearch_Request.SendForm = "Payments"
                       FrmSearch_Request.show
End If
End Sub

Private Sub TxtOther_Change()
'If Me.TxtModFlg.Text <> "R" Then
'TxtTotalsalary.Text = val(txtSalary.Text) + val(TxtSalEntitOther.Text) - val(TxtOther.Text) - val(TxtAdvance.Text) - val(TxtInsuranceValue.Text)
'XPTxtVal.Text = val(TxtTotalsalary) + val(txtSalaryVocation.Text) + val(Me.txtValueTickt.Text)
'End If
Calculte
End Sub

Private Sub TxtOther_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(TxtOther3.text) + val(Txtother.text), 2) > Round(val(TxtOther2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’»Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
Txtother.text = 0
TxtOther_Change
Txtother.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtPrePayd_Change(Index As Integer)
If DCboCashType.ListIndex = 7 Or DCboCashType.ListIndex = 1 Then
If val(TxtPrePayd(17).text) <> 0 Then
CalCulteVAT 1
XPTxtVal_Change
Else
CalcuteValue
End If
Else
If Me.TxtModFlg.text <> "R" And Me.TxtModFlg.text <> "" Then
Select Case Index
Case 13, 10, 12
If Round(val(TxtPrePayd(13).text) + val(TxtPrePayd(10).text), 2) < Round(val(TxtPrePayd(12).text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’·Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
TxtPrePayd(13).text = 0
' TxtPrePayd(13).SetFocus
Exit Sub
End If
Case 14, 15, 16
If Round(val(TxtPrePayd(15).text) + val(TxtPrePayd(16).text), 2) < Round(val(TxtPrePayd(14).text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’·Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
TxtPrePayd(13).text = 0
 TxtPrePayd(13).SetFocus
Exit Sub
End If
End Select

End If
CalculteTaoals
End If
End Sub

Private Sub TxtPrePayd_Click(Index As Integer)
TxtPrePayd_Change (13)
End Sub

Private Sub TxtPrePayd_GotFocus(Index As Integer)

If Index = 17 Then
If (DCboCashType.ListIndex = 7 Or DCboCashType.ListIndex = 1) Or Option3.value = True Then
TxtPrePayd(17).locked = False
Else
'TxtPrePayd(17).locked = True'salimcomment
End If
End If
'MsgBox TxtPrePayd(17).Enabled
End Sub



Private Sub TxtPrePayd_LostFocus(Index As Integer)
If Index = 17 Then
If val(TxtPrePayd(17).text) = 0 Then
'TxtPrePayd(17).Text = 0
Else
XPTxtVal_Change
End If
End If
End Sub

Private Sub txtSal_Change()
CalculteTaoals
End Sub
Sub CalculteTaoals()
If Me.TxtModFlg.text <> "R" Then
'If DCboCashType.ListIndex <> 4 Then Exit Sub
If DCboCashType.ListIndex <> 6 And DCboCashType.ListIndex <> 7 And DCboCashType.ListIndex <> 8 And DCboCashType.ListIndex <> 9 And DCboCashType.ListIndex <> 10 Then Exit Sub
txtNet2.text = val(txtTotal2.text) + val(txtSal2.text) + val(txtCustom.text) + val(txtTicketValue2.text) + val(TxtAddOther2.text)
txtNet.text = val(txtTotal.text) + val(txtSal.text) + val(txtCustom2.text) + val(txtTicketValue.text) + val(TxtAddOther.text)
TxtTotalDis2.text = val(TXTAdvanceTotal2.text) + val(TxtVlueVaction2.text) + val(TxtCash2.text) + val(TxtPrePayd(12).text) + val(TxtPrePayd(14).text)
TxtTotalDis.text = val(TXTAdvanceTotal.text) + val(TxtVlueVaction.text) + val(TxtCash.text) + val(TxtPrePayd(13).text) + val(TxtPrePayd(16).text)
TXTLastTotal.text = Round(val(txtNet.text) - val(TxtTotalDis.text), 10)
TXTLastTotal2.text = val(txtNet2.text) - val(TxtTotalDis2.text) - val(TxtTotlPaidEndSer.text)
TXTLastTotal.text = Round(val(TXTLastTotal.text), 2)
XPTxtVal.text = val(TXTLastTotal.text)
End If
End Sub

Private Sub txtSal_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(txtSal.text) + val(TxtPrePayd(1).text), 2) > Round(val(txtSal2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’·Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
txtSal.text = 0
txtSal_Change
txtSal.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtSalary_Change()
'If Me.TxtModFlg.Text <> "R" Then
'TxtTotalsalary.Text = val(txtSalary.Text) + val(TxtSalEntitOther.Text) - val(TxtOther.Text) - val(txtAdvance1.Text) - val(TxtInsuranceValue.Text)
'XPTxtVal.Text = val(TxtTotalsalary) + val(txtSalaryVocation.Text) + val(Me.txtValueTickt.Text)
'End If
Calculte
End Sub

Private Sub txtSalary_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(txtSalary3.text) + val(TxtSalary.text), 2) > Round(val(txtSalary2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’»Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
TxtSalary.text = 0
TxtSalary_Change
TxtSalary.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtSalaryVocation_Change()
Calculte
End Sub
Sub Calculte()
If Me.TxtModFlg.text <> "R" Then
TxtTotalsalary.text = val(TxtSalary.text) + val(TxtSalEntitOther.text) - val(Txtother.text) - val(txtAdvance1.text) - val(TxtInsuranceValue.text)
XPTxtVal.text = val(TxtTotalsalary) + val(txtSalaryVocation.text) + val(Me.txtValueTickt.text)
XPTxtVal.text = Round(val(XPTxtVal.text), 2)
End If
End Sub
Private Sub txtSalaryVocation_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(txtSalaryVocation3.text) + val(txtSalaryVocation.text), 2) > Round(val(txtSalaryVocation2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’»Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
txtSalaryVocation.text = 0
TxtSalaryVocation_Change
txtSalaryVocation.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtSalEntitOther_Change()
'If Me.TxtModFlg.Text <> "R" Then
'TxtTotalsalary.Text = val(txtSalary.Text) + val(TxtSalEntitOther.Text) - val(TxtOther.Text) - val(txtAdvance1.Text) - val(TxtInsuranceValue.Text)
'XPTxtVal.Text = val(TxtTotalsalary) + val(txtSalaryVocation.Text) + val(Me.txtValueTickt.Text)
'End If
Calculte
End Sub

Private Sub TxtSalEntitOther_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(TxtSalEntitOther3.text) + val(TxtSalEntitOther.text), 2) > Round(val(TxtSalEntitOther2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’»Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
TxtSalEntitOther.text = 0
TxtSalEntitOther_Change
TxtSalEntitOther.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub txtTicketValue_Change()
CalculteTaoals
End Sub

Private Sub txtTicketValue_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(txtTicketValue.text) + val(TxtPrePayd(3).text), 2) > Round(val(txtTicketValue2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’·Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
txtTicketValue.text = 0
txtTicketValue_Change
txtTicketValue.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtTotal_Change()
CalculteTaoals
End Sub

Private Sub txttotal_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(txtTotal.text) + val(TxtPrePayd(0).text), 2) > Round(val(txtTotal2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’·Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
txtTotal.text = 0
TxtTotal_Change
txtTotal.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtTotalDis_Change()
CalculteTaoals
End Sub

Private Sub txtTotalWithVat_LostFocus()
ClaCul
End Sub

Private Sub txtTotalWithVat_Validate(Cancel As Boolean)
    CalCulteVAT 0
End Sub

Private Sub txtTransferExpenses_Change()
XPTxtVal_Change
End Sub

Private Sub TxtTransferExpenses_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, txtTransferExpenses.text, 0)
    
End Sub

Private Sub TxtTransID_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If Me.TxtTransID.text <> "" Then
            If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                Me.TxtTransSerial.text = GetTransIDSerial(1, val(Me.TxtTransID.text))
            Else
                Me.TxtTransSerial.text = ""
            End If
        End If
    End If

End Sub

Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTransSerial.text, 1)
End Sub

Private Sub TxtValEndService_Change()
CalculteTaoals
End Sub

Private Sub TxtValueTickt_Change()
'If Me.TxtModFlg.Text <> "R" Then
'TxtTotalsalary.Text = val(txtSalary.Text) + val(TxtSalEntitOther.Text) - val(TxtOther.Text) - val(txtAdvance1.Text) - val(TxtInsuranceValue.Text)
'XPTxtVal.Text = val(TxtTotalsalary) + val(txtSalaryVocation.Text) + val(Me.txtValueTickt.Text)
'End If
Calculte
End Sub

Private Sub txtValueTickt_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(txtValueTickt3.text) + val(txtValueTickt.text), 2) > Round(val(txtValueTickt2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’»Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
txtValueTickt.text = 0
TxtValueTickt_Change
txtValueTickt.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub TxtVAt2_Change()
TxtVATValue.text = txtVat2.text
End Sub

Private Sub TxtVlueVaction_Change()
CalculteTaoals
End Sub
Private Sub TxtVATValue_Change()

If val(txtVat2.text) <> 0 Then
txtVat2.text = TxtVATValue.text
'XPTxtVal_Validate False
End If
End Sub

Private Sub TxtVlueVaction_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(TxtVlueVaction.text) + val(TxtPrePayd(8).text), 2) > Round(val(TxtVlueVaction2.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ œ›⁄ ﬁÌ„… «ﬂ»— „‰ «·ﬁÌ„… «·«’·Ì…"
Else
MsgBox "Can Not enter value Larher than total"
End If
TxtVlueVaction.text = 0
TxtVlueVaction_Change
TxtVlueVaction.SetFocus
Exit Sub
End If
End If
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VSFlexGrid1
Select Case VSFlexGrid1.ColKey(Col)
Case "TransPayedValueE"
    .TextMatrix(Row, .ColIndex("TransPayedValue")) = Round(val(.TextMatrix(Row, .ColIndex("TransPayedValueE"))) * val(.TextMatrix(Row, .ColIndex("Currency_rate"))), 3)
Case "TransPayedValue"
    .TextMatrix(Row, .ColIndex("TransPayedValueE")) = Round(val(.TextMatrix(Row, .ColIndex("TransPayedValue"))) / val(.TextMatrix(Row, .ColIndex("Currency_rate"))), 3)
End Select
End With
RelineBuy
RelineBu22

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid1
Select Case .ColKey(Col)
Case "TransPayedValue"
If .cell(flexcpChecked, Row, .ColIndex("payed")) = flexChecked Then
Cancel = False
Else
End If

Case "NoteSerial1"
Cancel = True
Case "too"
Cancel = True
Case "NoteDate"
Cancel = True
Case "branch_name"
Cancel = True
Case "Note_Value"
Cancel = True
Case "PayedValue"
Cancel = True
Case "RemainingValue"
Cancel = True
Case "NetValue"
Cancel = True

End Select
End With
End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
RelineProject22
RelineProject
End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid2
Select Case .ColKey(Col)
Case "TransPayedValue"
If .cell(flexcpChecked, Row, .ColIndex("payed")) = flexChecked Then
Cancel = False
Else
Cancel = True
End If
Case "Project_name"
Cancel = True
Case "NoteSerial1"
Cancel = True
Case "too"
Cancel = True
Case "NoteDate"
Cancel = True
Case "branch_name"
Cancel = True
Case "Note_Value"
Cancel = True
Case "PayedValue"
Cancel = True
Case "RemainingValue"
Cancel = True
Case "NetValue"
Cancel = True

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
Sub GetEmployee(Optional EmpID As Double = 0)
On Error Resume Next
If EmpID <> 0 Then
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "SELECT  * from TblEmployee"
sql = sql & "  Where (Emp_ID = " & EmpID & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
Me.TxtBeneficiaryBanck.text = IIf(IsNull(Rs7("BanckName").value), "", Rs7("BanckName").value)
'Me.TxtBeneficiaryAddress.text = IIf(IsNull(Rs7("Address").value), "", Rs7("Address").value)
Me.TxtBeneficiaryACNo.text = IIf(IsNull(Rs7("BankCard").value), "", Rs7("BankCard").value)
Me.TxtBenefiBanckAddress.text = IIf(IsNull(Rs7("BankIAddress").value), "", Rs7("BankIAddress").value)
Me.TxtBenefiBanckCode.text = IIf(IsNull(Rs7("BankCode").value), "", Rs7("BankCode").value)
If Not (IsNull(Rs7("NumEkama").value)) Then
If Rs7("NumEkama").value = "”⁄ÊœÌ" Or Rs7("NumEkama").value = "”⁄ÊœÏ" Or UCase$(Rs7("NumEkama").value) = "SAUDI" Then
TxtBenefNumIqama.text = IIf(IsNull(Rs7("NumPoket").value), "", Rs7("NumPoket").value)
BenefDateExpEqama.value = IIf(IsNull(Rs7("dateendpoketh").value), "", Rs7("dateendpoketh").value)
Else
TxtBenefNumIqama.text = IIf(IsNull(Rs7("NumEkama").value), "", Rs7("NumEkama").value)
BenefDateExpEqama.value = IIf(IsNull(Rs7("DateExpoekamaH").value), "", Rs7("DateExpoekamaH").value)
TxtBenefPlaceIqama.text = IIf(IsNull(Rs7("placeEkama").value), "", Rs7("placeEkama").value)
End If
End If
Dim X As Date

BenefBrithDate.value = IIf(IsNull(Rs7("DOB").value), Date, (Rs7("DOB").value))
TxtBenefTelephone.text = IIf(IsNull(Rs7("Emp_mobile").value), "", Rs7("Emp_mobile").value)
TxtBenefIBAN.text = IIf(IsNull(Rs7("BankIBan").value), "", Rs7("BankIBan").value)
Else
TxtBenefIBAN.text = ""
TxtBenefTelephone.text = ""
TXtBenefCountry.text = ""
TxtBenefGovernorate.text = ""
TxtBenefNumIqama.text = ""
Me.TxtBeneficiaryBanck.text = ""
Me.TxtBeneficiaryAddress.text = ""
Me.TxtBeneficiaryACNo.text = ""
Me.TxtBenefiBanckAddress.text = ""
Me.TxtBenefiBanckCode.text = ""
End If
End If
End Sub
Sub GetCustomer(Optional CusID As Double = 0)
If CusID <> 0 Then
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, "
sql = sql & "                       dbo.TblCustemers.Fullcode, dbo.TblCustemers.BankAddress, dbo.TblCustemers.BankIBAN, dbo.TblCustemers.BankCode, dbo.TblCustemers.IBAN,"
sql = sql & "                       dbo.TblCustemers.BankName, dbo.TblCustemers.BankAccount, dbo.TblCustemers.Address, dbo.TblCustemers.CustGID, dbo.TblCustemers.CountryID,"
sql = sql & "                       dbo.TblCountriesData.CountryName, dbo.TblCustemers.GovernmentID, dbo.TblCountriesGovernments.GovernmentName, dbo.TblCustemers.CityID,"
sql = sql & "                       dbo.TblCountriesGovernmentsCities.CityName"
sql = sql & "  FROM         dbo.TblCustemers LEFT OUTER JOIN"
sql = sql & "                       dbo.TblCountriesGovernmentsCities ON dbo.TblCustemers.CityID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblCountriesGovernments ON dbo.TblCustemers.GovernmentID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblCountriesData ON dbo.TblCustemers.CountryID = dbo.TblCountriesData.CountryID"
sql = sql & "  Where (dbo.TblCustemers.CusID = " & CusID & ")"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
Me.TxtBeneficiaryBanck.text = IIf(IsNull(Rs7("BankName").value), "", Rs7("BankName").value)
Me.TxtBeneficiaryAddress.text = IIf(IsNull(Rs7("Address").value), "", Rs7("Address").value)
Me.TxtBeneficiaryACNo.text = IIf(IsNull(Rs7("BankAccount").value), "", Rs7("BankAccount").value)
Me.TxtBenefiBanckAddress.text = IIf(IsNull(Rs7("BankAddress").value), "", Rs7("BankAddress").value)
Me.TxtBenefiBanckCode.text = IIf(IsNull(Rs7("BankCode").value), "", Rs7("BankCode").value)
TxtBenefNumIqama.text = IIf(IsNull(Rs7("CustGID").value), "", Rs7("CustGID").value)
TXtBenefCountry.text = IIf(IsNull(Rs7("CountryName").value), "", Rs7("CountryName").value)
TxtBenefGovernorate.text = IIf(IsNull(Rs7("GovernmentName").value), "", Rs7("GovernmentName").value)
TxtBenefTelephone.text = IIf(IsNull(Rs7("Cus_Phone").value), "", Rs7("Cus_Phone").value)
TxtBenefIBAN.text = IIf(IsNull(Rs7("BankIBAN").value), "", Rs7("BankIBAN").value)
Else
TxtBenefIBAN.text = ""
TxtBenefTelephone.text = ""
TXtBenefCountry.text = ""
TxtBenefGovernorate.text = ""
TxtBenefNumIqama.text = ""
Me.TxtBeneficiaryBanck.text = ""
Me.TxtBeneficiaryAddress.text = ""
Me.TxtBeneficiaryACNo.text = ""
Me.TxtBenefiBanckAddress.text = ""
Me.TxtBenefiBanckCode.text = ""
End If
End If
End Sub



Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim RsDev As ADODB.Recordset
    Dim i As Integer
  
     On Error GoTo ErrTrap
     ReloadContracR
    Option4.value = False
    Option5.value = False
    Option6.value = False
    Option7.value = False
    Fra(2).Visible = False
    lbl(47).Visible = False
        TxtAdvance.Visible = False

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
   
    TxtVATValue = txtVat2
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


    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    End If
If Not (IsNull(rs("LockedInterval").value)) Then
If rs("LockedInterval").value = True Then
Cmd(1).Enabled = False
Cmd(4).Enabled = False
Else
Cmd(1).Enabled = True
Cmd(4).Enabled = True
End If
Else
Cmd(1).Enabled = True
Cmd(4).Enabled = True
End If
''///////
TxtCurrencyRate.text = IIf(IsNull(rs("Rate").value), 1, rs("Rate").value)
DcbCurrency.BoundText = IIf(IsNull(rs("CurrncyID").value), "", rs("CurrncyID").value)
XPTxtValE.text = IIf(IsNull(rs("Note_ValueE").value), 0, rs("Note_ValueE").value)

DcbEmpBranch.BoundText = IIf(IsNull(rs("EmpBranchID").value), "", rs("EmpBranchID").value)
DcbAccount.BoundText = IIf(IsNull(rs("AccountPaym").value), "", rs("AccountPaym").value)
TxtPrePayd(0).text = IIf(IsNull(rs("total3").value), "", rs("total3").value)
TxtPrePayd(1).text = IIf(IsNull(rs("Sal3").value), "", rs("Sal3").value)
TxtPrePayd(2).text = IIf(IsNull(rs("Custom3").value), "", rs("Custom3").value)
TxtPrePayd(3).text = IIf(IsNull(rs("TicketValue3").value), "", rs("TicketValue3").value)
TxtPrePayd(4).text = IIf(IsNull(rs("AddOther3").value), "", rs("AddOther3").value)
TxtPrePayd(5).text = IIf(IsNull(rs("net3").value), "", rs("net3").value)
TxtPrePayd(6).text = IIf(IsNull(rs("ValEndService3").value), "", rs("ValEndService3").value)
TxtPrePayd(7).text = IIf(IsNull(rs("AdvanceTotal3").value), "", rs("AdvanceTotal3").value)
TxtPrePayd(8).text = IIf(IsNull(rs("VlueVaction3").value), "", rs("VlueVaction3").value)
TxtPrePayd(9).text = IIf(IsNull(rs("Cash3").value), "", rs("Cash3").value)
TxtPrePayd(10).text = IIf(IsNull(rs("Discounts3").value), "", rs("Discounts3").value)
TxtPrePayd(11).text = IIf(IsNull(rs("TotalDis3").value), "", rs("TotalDis3").value)
TxtPrePayd(15).text = IIf(IsNull(rs("DisSalaryPayed").value), "", rs("DisSalaryPayed").value)
''//////
Option1.value = False
Option3.value = False

If IsNull(rs("NCashingType").value) Then

Else
        If rs("NCashingType").value = 1 Then
               Option1.value = True
        ElseIf rs("NCashingType").value = 3 Then
             Option3.value = True
        End If
End If

TxtPrePayd(17).text = IIf(IsNull(rs("PreVAT").value), 0, rs("PreVAT").value)
If Not IsNull(rs("IncludVAT").value) Then
If (rs("IncludVAT").value) = 1 Then
IncludVAT.value = vbChecked
Else
IncludVAT.value = vbUnchecked
End If
Else
IncludVAT.value = vbUnchecked
End If
    TxtTotlPaidEndSer.text = IIf(IsNull(rs("TotlPaidEndSer").value), "", rs("TotlPaidEndSer").value)
    TxtPrePayd(13).text = IIf(IsNull(rs("Discounts").value), "", rs("Discounts").value)
    TxtPrePayd(12).text = IIf(IsNull(rs("Discounts2").value), "", rs("Discounts2").value)
    TxtPrePayd(14).text = IIf(IsNull(rs("DisSalary2").value), "", rs("DisSalary2").value)
    TxtPrePayd(16).text = IIf(IsNull(rs("DisSalary").value), "", rs("DisSalary").value)
    
    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    DcbDepartment.BoundText = IIf(IsNull(rs("DeptID").value), 0, rs("DeptID").value)
    
     Me.TxtNumIqama.text = IIf(IsNull(rs("NumIqama").value), "", rs("NumIqama").value)
     Me.TxtTelephone.text = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
     
     Me.TxtOrder.text = IIf(IsNull(rs("OrderIDD").value), "", rs("OrderIDD").value)
     Me.TxtAdvance.text = IIf(IsNull(rs("AdvanceIDD").value), "", rs("AdvanceIDD").value)
      DCPreFix.text = IIf(IsNull(rs("Prefix").value), "", rs("Prefix").value)

    EmpIDD = IIf(IsNull(rs("EmpIDD").value), 0, rs("EmpIDD").value)
    Me.TXT_order_no.text = IIf(IsNull(rs("Order_no").value), "", rs("Order_no").value)
    Me.Text1.text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    XPTxtID1.text = IIf(IsNull(rs("AdvanceID").value), "", (rs("AdvanceID").value))
    txtTransferExpenses.text = IIf(IsNull(rs("TransferExpenses").value), "", (rs("TransferExpenses").value))
    TxtManulaNO.text = IIf(IsNull(rs("ManulaNO").value), "", rs("ManulaNO").value)

    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
''''////////////
If Not IsNull(rs("Fina_Pruch").value) Then
If rs("Fina_Pruch").value = 1 Then
Option8(1).value = True
Else
Option8(0).value = True
End If
Else
Option8(0).value = True
End If

Me.dcDur.BoundText = IIf(IsNull(rs("Year_Req").value), 0, rs("Year_Req").value)
Me.dcMontth.BoundText = IIf(IsNull(rs("Month_Req").value), 0, rs("Month_Req").value)
Me.DcbBrReq.BoundText = IIf(IsNull(rs("Brnch_Req").value), 0, rs("Brnch_Req").value)
Me.TxtBeneficiaryBanck.text = IIf(IsNull(rs("BeneficiaryBanck").value), "", rs("BeneficiaryBanck").value)
Me.TxtBeneficiaryAddress.text = IIf(IsNull(rs("BeneficiaryAddress").value), "", rs("BeneficiaryAddress").value)
Me.TxtBeneficiaryACNo.text = IIf(IsNull(rs("BeneficiaryACNo").value), "", rs("BeneficiaryACNo").value)
Me.TxtBenefiBanckAddress.text = IIf(IsNull(rs("BenefiBanckAddress").value), "", rs("BenefiBanckAddress").value)
Me.TxtBenefiBanckCode.text = IIf(IsNull(rs("BenefiBanckCode").value), "", rs("BenefiBanckCode").value)
Me.TxtRemitterName.text = IIf(IsNull(rs("RemitterName").value), "", rs("RemitterName").value)
Me.TxtReportName.text = IIf(IsNull(rs("ReportName").value), "", rs("ReportName").value)
Me.TxtBenefTelephone.text = IIf(IsNull(rs("BenefTelephone").value), "", rs("BenefTelephone").value)
'''////////
Me.TxtBenefPlaceBrith.text = IIf(IsNull(rs("BenefPlaceBrith").value), "", rs("BenefPlaceBrith").value)
BenefBrithDate.value = IIf(IsNull(rs("BenefBrithDate").value), Date, rs("BenefBrithDate").value)
Me.BenefDateExpEqama.value = IIf(IsNull(rs("BenefDateExpEqama").value), "", rs("BenefDateExpEqama").value)
Me.TxtKafelAddress.text = IIf(IsNull(rs("KafelAddress").value), "", rs("KafelAddress").value)
Me.TxtBenefPlaceIqama.text = IIf(IsNull(rs("BenefPlaceIqama").value), "", rs("BenefPlaceIqama").value)
Me.TxtBenefIBAN.text = IIf(IsNull(rs("BenefIBAN").value), "", rs("BenefIBAN").value)
'''///////
Me.TxtBenefNumIqama.text = IIf(IsNull(rs("BenefNumIqama").value), "", rs("BenefNumIqama").value)
Me.TxtTicktConract.text = IIf(IsNull(rs("TicktConract").value), 0, rs("TicktConract").value)
Me.TXtBenefCountry.text = IIf(IsNull(rs("BenefCountry").value), "", rs("BenefCountry").value)
Me.TxtBenefCity.text = IIf(IsNull(rs("BenefCity").value), "", rs("BenefCity").value)
Me.TxtBenefStreet.text = IIf(IsNull(rs("BenefStreet").value), "", rs("BenefStreet").value)
Me.TxtBenefGovernorate.text = IIf(IsNull(rs("BenefGovernorate").value), "", rs("BenefGovernorate").value)
Me.TxtCountry.text = IIf(IsNull(rs("Country").value), "", rs("Country").value)
Me.TxtGovernorate.text = IIf(IsNull(rs("Governorate").value), "", rs("Governorate").value)
Me.TxtCity.text = IIf(IsNull(rs("City").value), "", rs("City").value)
Me.TxtStreet.text = IIf(IsNull(rs("Street").value), "", rs("Street").value)
Me.TxtAdress2.text = IIf(IsNull(rs("Adress2").value), "", rs("Adress2").value)
Me.TxtKafelName.text = IIf(IsNull(rs("KafelName").value), "", rs("KafelName").value)
Me.TxtKafeltEL.text = IIf(IsNull(rs("KafeltEL").value), "", rs("KafeltEL").value)
Me.TxtInsuranceValue.text = IIf(IsNull(rs("InsuranceValue").value), 0, rs("InsuranceValue").value)
    '**********************************12 12 2015
       Me.CboYear1.ListIndex = IIf(IsNull(rs("PayrollYear").value), 0, rs("PayrollYear").value)
       Me.CmbMonth1.ListIndex = IIf(IsNull(rs("PayrollMonth").value), 0, rs("PayrollMonth").value - 1)
      If Me.CboYear1.ListIndex = -1 Then
      CboYear1.text = year(Date)
      End If
    '   Me.CmbMonth1.ListIndex = IIf(IsNull(rs("PayrollMonth").value), 0, rs("PayrollMonth").value)
''//////////
Me.txtTotal2.text = IIf(IsNull(rs("Total2").value), 0, rs("Total2").value)
Me.txtSal2.text = IIf(IsNull(rs("Sal2").value), 0, rs("Sal2").value)
Me.txtNet2.text = IIf(IsNull(rs("net2").value), 0, rs("net2").value)
Me.TxtVlueVaction2.text = IIf(IsNull(rs("VlueVaction2").value), 0, rs("VlueVaction2").value)
Me.txtTicketValue2.text = IIf(IsNull(rs("TicketValue2").value), 0, rs("TicketValue2").value)
Me.txtCustom2.text = IIf(IsNull(rs("Custom2").value), 0, rs("Custom2").value)
Me.TXTAdvanceTotal2.text = IIf(IsNull(rs("AdvanceTotal2").value), 0, rs("AdvanceTotal2").value)
Me.TxtCash2.text = IIf(IsNull(rs("Cash2").value), 0, rs("Cash2").value)
Me.txtSalary2.text = IIf(IsNull(rs("Salary2").value), 0, rs("Salary2").value)
Me.TxtSalEntitOther2.text = IIf(IsNull(rs("SalEntitOther2").value), 0, rs("SalEntitOther2").value)
Me.TxtOther2.text = IIf(IsNull(rs("Other2").value), 0, rs("Other2").value)
Me.txtAdvance12.text = IIf(IsNull(rs("Advance12").value), 0, rs("Advance12").value)
Me.TxtInsuranceValue2.text = IIf(IsNull(rs("InsuranceValue2").value), 0, rs("InsuranceValue2").value)
Me.TXTLastTotal2.text = IIf(IsNull(rs("LastTotal2").value), 0, rs("LastTotal2").value)
Me.txtSalaryVocation2.text = IIf(IsNull(rs("SalaryVocation2").value), 0, rs("SalaryVocation2").value)
Me.txtValueTickt2.text = IIf(IsNull(rs("ValueTickt2").value), 0, rs("ValueTickt2").value)
Me.TxtTotalsalary2.text = IIf(IsNull(rs("Totalsalary2").value), 0, rs("Totalsalary2").value)
Me.TxtCusTiket2.text = IIf(IsNull(rs("CusTiket2").value), 0, rs("CusTiket2").value)
Me.TxtCusTiket.text = IIf(IsNull(rs("CusTiket").value), 0, rs("CusTiket").value)
Me.TxtAddOther.text = IIf(IsNull(rs("AddOther").value), 0, rs("AddOther").value)
Me.TxtAddOther2.text = IIf(IsNull(rs("AddOther2").value), 0, rs("AddOther2").value)
Me.TxtValEndService2.text = IIf(IsNull(rs("ValEndService2").value), 0, rs("ValEndService2").value)
Me.TxtValEndService.text = IIf(IsNull(rs("ValEndService").value), 0, rs("ValEndService").value)
Me.TxtTotalDis2.text = IIf(IsNull(rs("TotalDis2").value), 0, rs("TotalDis2").value)
Me.TxtTotalDis.text = IIf(IsNull(rs("TotalDis").value), 0, rs("TotalDis").value)
Me.TXTAdvanceTotal.text = IIf(IsNull(rs("AdvanceTotal").value), 0, rs("AdvanceTotal").value)
Me.TxtVlueVaction.text = IIf(IsNull(rs("VlueVaction").value), 0, rs("VlueVaction").value)
Me.TxtCash.text = IIf(IsNull(rs("Cash").value), 0, rs("Cash").value)
Me.TXTLastTotal2.text = IIf(IsNull(rs("LastTotal2").value), 0, rs("LastTotal2").value)
Me.TXTLastTotal.text = IIf(IsNull(rs("LastTotal").value), 0, rs("LastTotal").value)
Me.TxtCusTiket.text = IIf(IsNull(rs("CusTiket").value), 0, rs("CusTiket").value)
Me.txtCustom.text = IIf(IsNull(rs("Custom").value), 0, rs("Custom").value)
Me.txtTicketValue.text = IIf(IsNull(rs("TicketValue").value), 0, rs("TicketValue").value)
Me.txtSal.text = IIf(IsNull(rs("Sal").value), 0, rs("Sal").value)
Me.txtTotal.text = IIf(IsNull(rs("total").value), 0, rs("total").value)
Me.DcbEmpEndService.BoundText = IIf(IsNull(rs("EmpIDEnd").value), 0, rs("EmpIDEnd").value)
Me.DcbBranchEndServ.BoundText = IIf(IsNull(rs("BrnchIDEnd").value), 0, rs("BrnchIDEnd").value)
Me.DcboEmpName.BoundText = IIf(IsNull(rs("EmpIDVac").value), 0, rs("EmpIDVac").value)
Me.dcBranch1.BoundText = IIf(IsNull(rs("BrnchIDVac").value), 0, rs("BrnchIDVac").value)
Me.TxtSalary.text = IIf(IsNull(rs("Salary33").value), 0, rs("Salary33").value)
Me.TxtSalEntitOther.text = IIf(IsNull(rs("SalEntitOther").value), 0, rs("SalEntitOther").value)
Me.Txtother.text = IIf(IsNull(rs("Other").value), 0, rs("Other").value)
Me.txtAdvance1.text = IIf(IsNull(rs("Advance1").value), 0, rs("Advance1").value)
Me.TxtInsuranceValue.text = IIf(IsNull(rs("InsuranceValue").value), 0, rs("InsuranceValue").value)
Me.txtSalaryVocation.text = IIf(IsNull(rs("SalaryVocation").value), 0, rs("SalaryVocation").value)
Me.txtValueTickt.text = IIf(IsNull(rs("ValueTickt").value), 0, rs("ValueTickt").value)

Me.txtSalary3.text = IIf(IsNull(rs("Salary3").value), 0, rs("Salary3").value)
Me.TxtSalEntitOther3.text = IIf(IsNull(rs("SalEntitOther3").value), 0, rs("SalEntitOther3").value)
Me.txtSalaryVocation3.text = IIf(IsNull(rs("SalaryVocation3").value), 0, rs("SalaryVocation3").value)
Me.TxtOther3.text = IIf(IsNull(rs("Other3").value), 0, rs("Other3").value)
Me.txtValueTickt3.text = IIf(IsNull(rs("ValueTickt3").value), 0, rs("ValueTickt3").value)
Me.txtAdvance13.text = IIf(IsNull(rs("Advance13").value), 0, rs("Advance13").value)
Me.TxtInsuranceValue3.text = IIf(IsNull(rs("InsuranceValue3").value), 0, rs("InsuranceValue3").value)

''/////
       Me.empDes.text = IIf(IsNull(rs("empDes").value), "", rs("empDes").value)
       Me.empDes1.text = IIf(IsNull(rs("empDes1").value), "", rs("empDes1").value)
 Me.PayDes.text = IIf(IsNull(rs("PayDes").value), "", rs("PayDes").value)
      '**********************************12 12 2015
      Me.TxtOrderSuppler.text = IIf(IsNull(rs("TxtOrderSuppler").value), "", rs("TxtOrderSuppler").value)
       Me.TxtNoSupplerDes.text = IIf(IsNull(rs("TxtNoSupplerDes").value), "", rs("TxtNoSupplerDes").value)
     
     
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(45).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    txt_general_des.text = IIf(IsNull(rs("general_des_notes").value), "", rs("general_des_notes").value)

    txtperson.text = IIf(IsNull(rs("person").value), "", rs("person").value)

   ' XPTxtVal.Text = IIf(IsNull(rs("Note_Value2").value), IIf(IsNull(rs("Note_Value").value), 0, (rs("Note_Value").value)), (rs("Note_Value2").value)) - IIf(IsNull(rs("PreVAT").value), 0, (rs("PreVAT").value))
    XPTxtVal.text = IIf(IsNull(rs("Note_Value2").value), IIf(IsNull(rs("Note_Value").value), 0, (rs("Note_Value").value)), (rs("Note_Value2").value))
    dcproject.BoundText = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
    XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))

    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    Txt_DateHigri.value = IIf(IsNull(rs("NoteDateH").value), ToHijriDate(XPDtbTrans.value), rs("NoteDateH").value)

    DCboCashType.ListIndex = IIf(IsNull(rs("CashingType").value), -1, rs("CashingType").value)
   
TxtDue.text = IIf(IsNull(rs("Due").value), "", rs("Due").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
''//
   Me.dcproject1.BoundText = IIf(IsNull(rs("ProjectID").value), "", rs("ProjectID").value)
   Me.Dcterm1.BoundText = IIf(IsNull(rs("Pand").value), "", rs("Pand").value)
   Me.dcopr.BoundText = IIf(IsNull(rs("Oper").value), "", rs("Oper").value)
'''/
    If IsNull(rs("NoteCashingType").value) Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
    ElseIf rs("NoteCashingType").value = 1 Then
        Me.CboPayMentType.ListIndex = 1
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
 
    ElseIf rs("NoteCashingType").value = 2 Then
        Me.CboPayMentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    ElseIf rs("NoteCashingType").value = 3 Then
        Me.CboPayMentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        ElseIf rs("NoteCashingType").value = 4 Then
        Me.CboPayMentType.ListIndex = 4
    ElseIf rs("NoteCashingType").value = 5 Then
        Me.CboPayMentType.ListIndex = 5
    End If

    If DCboCashType.ListIndex = 3 Then
        DBCboClientName.BoundText = IIf(IsNull(rs("ProjectMainID").value), 0, rs("ProjectMainID").value)

    ElseIf DCboCashType.ListIndex = 4 Then

        If IsNull(rs("salary_or_advance").value) Then
            Option4.value = False: Option5.value = False
        ElseIf (rs("salary_or_advance").value) = 0 Then
            Option4.value = True
            Option4_Click
        ElseIf (rs("salary_or_advance").value) = 1 Then
            Option5.value = True
            Option5_Click
        ElseIf (rs("salary_or_advance").value) = 2 Then
            Option6.value = True
            Option6_Click
        ElseIf (rs("salary_or_advance").value) = 3 Then
            Option7.value = True
            Option7_Click
        End If

        

        DBCboClientName.BoundText = IIf(IsNull(rs("EmpAccountCode").value), 0, rs("EmpAccountCode").value)

    ElseIf DCboCashType.ListIndex = 5 Or DCboCashType.ListIndex = 6 Then
        DBCboClientName.BoundText = IIf(IsNull(rs("BTCashAccountcode").value), 0, rs("BTCashAccountcode").value)
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
    End If
 If DCboCashType.ListIndex = 12 Then
    TxtEndService.text = IIf(IsNull(rs("VATVowalNo").value), "", rs("VATVowalNo").value)
    Else
    TxtEndService.text = IIf(IsNull(rs("TxtEndService").value), "", rs("TxtEndService").value)
    End If
    '---------------------------------------------------------------------------
DcbContractor.BoundText = IIf(IsNull(rs("ContractorID").value), "", rs("ContractorID").value)
    RetriveBillVendorData
    RetriveBillBuyData
    RetriveBillProjectData
CboPayMentType_Change
    '-----------------------------------------------------------------------------
    If Not IsNull(rs("Transaction_ID").value) Then
        Me.ChkTrans.value = vbChecked
        'Me.ChkTrans.Enabled = True
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select * From Transactions Where Transaction_ID=" & rs("Transaction_ID").value
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            Me.TxtTransID.text = RsTemp("Transaction_ID").value
            Me.TxtTransSerial.text = IIf(IsNull(RsTemp("Transaction_Serial").value), "", RsTemp("Transaction_Serial").value)

            If Not (IsNull(RsTemp("Transaction_Type").value)) Then
                If RsTemp("Transaction_Type").value = 9 Then
                    Me.CboTrans.ListIndex = 1
                ElseIf RsTemp("Transaction_Type").value = 1 Then
                    Me.CboTrans.ListIndex = 0
                End If
            End If
        End If

    Else
        Me.ChkTrans.value = vbUnchecked
        Me.CboTrans.ListIndex = -1
        Me.TxtTransID.text = ""
        Me.TxtTransSerial.text = ""
    End If
If IsNull(rs("subcontractType").value) Then
subContOpt(2).value = True
ElseIf (rs("subcontractType").value) = 0 Then
subContOpt(0).value = True

ElseIf (rs("subcontractType").value) = 1 Then
subContOpt(1).value = True

ElseIf (rs("subcontractType").value) = 2 Then
subContOpt(2).value = True


End If
txtTradingContractID = IIf(IsNull(rs("TradingContractID").value), 0, rs("TradingContractID").value)
XPTxtValE.text = Format(XPTxtValE.text, "#,##0.00")
          XPTxtVal.text = Format(XPTxtVal.text, "#,##0.00")
 
    '-----------------------------------------------------------------------------
  If DCboCashType.ListIndex <> 6 And DCboCashType.ListIndex <> 7 And DCboCashType.ListIndex <> 8 And DCboCashType.ListIndex <> 9 And DCboCashType.ListIndex <> 10 Then
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.XPTxtID.text)
        StrSQL = StrSQL + " Order By DEV_ID_Line_No desc "
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (RsDev.BOF Or rs.EOF) Then
                        Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
                        Me.lbl(11).Caption = RsDev("Account_Interval_ID").value
                        RsDev.MoveFirst
                        
                                    For i = 1 To RsDev.RecordCount
                        
                                        If RsDev("Credit_Or_Debit").value = 0 Then
                                            Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                                        ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                                            Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                                        End If
                        
                                        RsDev.MoveNext
                                    Next i
            
                    End If
    End If

    Else
    If Me.DcbAccount.BoundText = "" Then
    
        Me.DcboDebitSide.BoundText = ""
         Me.DcboCreditSide.BoundText = ""
         

         If DcboBox.BoundText <> "" And (CboPayMentType.ListIndex < 4) Then
         'If CboPayMentType.ListIndex = 4 Then
         'Me.DcboCreditSide.BoundText = Me.DcbAccount.BoundText
         'Else
         Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
        ' End If
       '  DcboBox_Change
         End If
       End If
                  If DcboBankName.BoundText <> "" Then
          
          
              If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If 1 = 1 Then
        '    Me.DcboCreditSide.BoundText = "a2a3a2"
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If SystemOptions.banks_Accounts3 = True Then
            Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code2")
        Else
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
        
        If CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then
                     
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If
        
        'Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value

    End If


         End If
         
         
    End If

    
    '-----------------------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
     txtVat2 = TxtPrePayd(17)
     CalCulteVAT 3
     
     txtTotalWithVat.text = IIf(IsNull(rs("TotalNotesValue").value), val(txtVat2) + val(Format((XPTxtVal.text), "###.00")), rs("TotalNotesValue").value)
           With FG
 
                    Me.LblTotalV.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("PartValue"), .rows - 1, .ColIndex("PartValue"))
              
                End With
                
                
    Exit Sub
ErrTrap:
End Sub
Function CheckAdvanecPayed() As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from TblEmpAdvancePayedDet where AdvanceID=" & val(XPTxtID1.text) & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckAdvanecPayed = True
Else
CheckAdvanecPayed = False
End If
End Function

 
Sub CalCuteCurrencyE()
XPTxtVal.text = Round(val(Format(XPTxtValE.text, "###.00")) * val(TxtCurrencyRate.text), 2)
If val(TxtCurrencyRate.text) = 0 Then
TxtCurrencyRate.text = ""
End If
XPTxtVal.text = Round(val(Format(XPTxtValE.text, "###.00")) * val(TxtCurrencyRate.text), 2)

XPTxtVal.text = Format(XPTxtVal.text, "#,##0.00")

If val(XPTxtValE.text) = 0 Then XPTxtVal.text = "": Exit Sub

End Sub
Function GetBrnchCustomer(Optional CusID As Double) As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     BranchId"
sql = sql & " From dbo.TblCustemers"
sql = sql & " Where (CusID = " & CusID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetBrnchCustomer = IIf(IsNull(rs2("BranchId").value), val(Me.dcBranch.BoundText), rs2("BranchId").value)
Else
GetBrnchCustomer = val(Me.dcBranch.BoundText)
End If
If GetBrnchCustomer = 0 Then
GetBrnchCustomer = val(Me.dcBranch.BoundText)
End If
End Function
Private Sub SaveData()
    Dim Msg As String
      Dim total_value As Double
      Dim total_valuee As Double
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim StrTemp As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
    Dim sql As String
    Dim Percetage As Double
    Dim AccountVATCreit As String
'      On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
If DCboCashType.ListIndex = 12 Then
            Account_Code_dynamic = get_account_code_branch(145, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "No Branch Created", vbCritical
                End If

                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ÂÌ∆… «·“ﬂ«…", vbCritical
                    Else
                        MsgBox "Please Select Account VAT ", vbCritical
                    End If

                    GoTo ErrTrap
         
                End If
            End If
End If
        If CboPayMentType.ListIndex = 2 And val(Me.txtTransferExpenses.text) > 0 Then
            Account_Code_dynamic = get_account_code_branch(52, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "No Branch Created", vbCritical
                End If

                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»       „’—Ê›«   »‰ﬂÌ…  ›Ì  ‘«‘… —»ÿ «·Õ”«»«   ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "The bank Commisiion Account in this Branch is not specific", vbCritical
                    End If

                    GoTo ErrTrap
         
                End If
            End If
        
        End If
        
        If DCboCashType.ListIndex = -1 Then
                        If SystemOptions.UserInterface = ArabicInterface Then

            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·„œ›Ê⁄«  "
            
            Else
            Msg = "Define Payment Type Firstly"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboCashType.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If

        If DBCboClientName.text = "" And DCboCashType.ListIndex <> 6 And DCboCashType.ListIndex <> 12 And DCboCashType.ListIndex <> 10 And DCboCashType.ListIndex <> 7 And DCboCashType.ListIndex <> 11 And DCboCashType.ListIndex <> 8 And DCboCashType.ListIndex <> 9 Then
      If SystemOptions.UserInterface = ArabicInterface Then
            
            Msg = "ÌÃ» «Œ Ì«—«·«”„"
Else
   Msg = "Please Select Name"
End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DBCboClientName.SetFocus
        Sendkeys "{F4}"
            Exit Sub
        End If

        If XPTxtVal.text = "" Then
              If SystemOptions.UserInterface = ArabicInterface Then

            Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·„œ›Ê⁄«  "
            Else
            Msg = "Please Enter Value"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtVal.SetFocus
            Exit Sub
        End If

        If Not IsNumeric(XPTxtVal.text) Then
                      If SystemOptions.UserInterface = ArabicInterface Then

            Msg = "ﬁÌ„… «·„œ›Ê⁄«  ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            Else
            Msg = "Payment nust be Numeric"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtVal.SetFocus
            SelectText XPTxtVal
            Exit Sub
        End If

        If Me.CboPayMentType.ListIndex = -1 Then
                              If SystemOptions.UserInterface = ArabicInterface Then

            Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
            Else
            Msg = "Select Payment Method ...!!!"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPayMentType.SetFocus
            Exit Sub
        End If
'''//

        If Me.CboPayMentType.ListIndex = 0 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
                               If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
                Else
                Msg = "Select Box Firstly..!!"
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
                Msg = "Select Bank Firstly...!!"
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
             Msg = "Enter Cheque No Firstly...!!"
             End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If

            '      If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
            '          Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            '          MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '          DtpChequeDueDate.SetFocus
            '          SendKeys "{F4}"
            '          Exit Sub
            '      End If
        ElseIf Me.CboPayMentType.ListIndex = 4 Or Me.CboPayMentType.ListIndex = 5 Then
        If DcbAccount.BoundText = "" Or DcbAccount.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«»"
        Else
        MsgBox "Please Select Account"
        End If
        DcbAccount.SetFocus
        Exit Sub
        End If
        ElseIf Me.CboPayMentType.ListIndex = 2 Then

            If Me.DcboBankName.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ   «·»‰ﬂ...!!"
             Else
             Msg = " Specify Bank    ...!!"
             End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
               Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ —ﬁ„ «·ÕÊ«·Â...!!"
             Else
             Msg = " Define Transfer No#    ...!!"
             End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If
  
        End If
    
        If Me.TxtModFlg.text = "N" Then
            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), XPDtbTrans.value) = False Then
                        Exit Sub
                    End If
                End If
            End If

        ElseIf Me.TxtModFlg.text = "E" Then

            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), XPDtbTrans.value, , , val(Me.XPTxtID.text)) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        If Me.ChkTrans.value = vbChecked Then
            If Me.CboTrans.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·›« Ê—…..!!!"
            Else
            Msg = "Select Invoice firstly..!!!"
            End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                CboTrans.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim(Me.TxtTransSerial.text) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡ ≈œŒ«· —ﬁ„ «·›« Ê—…..!!!"
            Else
            Msg = " Enter Invoice #      ..!!!"
            End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Sub
            Else

                If Me.CboTrans.ListIndex = 0 Then
                    If Me.TxtTransID.text = "" Then
                        StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 1, Me.DBCboClientName.BoundText)
                    Else
                        StrTemp = Me.TxtTransID.text
                    End If

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 1 Then
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), 9)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If
    my_branch = val(Me.dcBranch.BoundText)
        If TxtNoteSerial.text = "" Then
            If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                     Else
                     MsgBox "GL NO Exceede ": Exit Sub
                     End If
            Else
                       
                If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                                  MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                         Else
                               MsgBox "Cant Create GL enter number Manually  ": Exit Sub
                         End If
                Else
                    TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
                End If
            End If
        End If
        If val(TxtCurrencyRate.text) = 0 Then
        TxtCurrencyRate.text = 1
        End If
        If TxtNoteSerial1.text = "" Then
            If Voucher_coding(val(my_branch), XPDtbTrans.value, 4, 5, , , DCPreFix.text) = "error" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ œ›⁄ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
              Else
              MsgBox "  Voucher Exceed Coding   ": Exit Sub
              End If
              
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbTrans.value, 4, 5, , , DCPreFix.text) = "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then

                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                    MsgBox "  Enter Number manually   ": Exit Sub
                    End If
                Else
                    TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 4, 5, , , DCPreFix.text)
                End If
            End If
        End If
        Dim i As Integer
   If val(DCboCashType.ListIndex) = 1 Then
   With GRID1
   If .rows >= 2 Then
   For i = 1 To .rows - 1
   If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And .cell(flexcpChecked, i, .ColIndex("haveqest")) = flexChecked And .TextMatrix(i, .ColIndex("StrQest")) = "" Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ   ÌÃ» «‰  ”œœ « ﬁ”«ÿ «·›« Ê—… —ﬁ„" & .TextMatrix(i, .ColIndex("NoteSerial1"))
   Else
   MsgBox "Can Not Save you must pay installments invoice number " & .TextMatrix(i, .ColIndex("NoteSerial1"))
   End If
   Exit Sub
   End If
   Next i
   End If
   End With
   End If

        Cn.BeginTrans
        BeginTrans = True
        '   If Option4.value = True Then
        '      txt_general_des.text = Option4.Caption
        '    ElseIf Option5.value = True Then
        '      txt_general_des.text = Option5.Caption
        '      ElseIf Option6.value = True Then
        '      txt_general_des.text = Option6.Caption
        '
        '        ElseIf Option7.value = True Then
        '      txt_general_des.text = Option7.Caption
        '
        '      End If

        If TxtModFlg.text = "N" Then
            XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
            '  Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=5"))
            rs.AddNew
            rs("NoteID").value = val(XPTxtID.text)
            XPTxtID.text = IIf(IsNull(rs("NoteID").value), 0, rs("NoteID").value)
            Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
        ElseIf TxtModFlg.text = "E" Then
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute "Delete from TblSalaryNotesPayment where TransID=" & val(XPTxtID.text) & ""
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
    StrSQL = "Delete From TblEmpAdvance Where AdvanceID=" & val(Me.XPTxtID1.text)

    Cn.Execute StrSQL, , adExecuteNoRecords
        
        End If
         XPTxtVal.text = Format(XPTxtVal.text, "###.00")
         XPTxtValE.text = Format(XPTxtValE.text, "###.00")
    '''///////////
    rs("Note_ValueE").value = IIf(val(XPTxtValE.text) = 0, Null, val(XPTxtValE.text))
    rs("Rate").value = IIf(val(TxtCurrencyRate.text) = 0, Null, val(TxtCurrencyRate.text))
    rs("CurrncyID").value = IIf(Trim(Me.DcbCurrency.BoundText) = "", Null, val(DcbCurrency.BoundText))
    rs("ContractorID").value = IIf(Trim(DcbContractor.BoundText) = "", Null, val(DcbContractor.BoundText))
    rs("EmpBranchID").value = IIf(Trim(DcbEmpBranch.BoundText) = "", Null, val(DcbEmpBranch.BoundText))
    rs("AccountPaym").value = IIf(Trim(DcbAccount.BoundText) = "", Null, DcbAccount.BoundText)
    rs("total3").value = IIf(Trim(Me.TxtPrePayd(0).text) = "", Null, val(Me.TxtPrePayd(0).text))
    rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, val(XPTxtVal.text))
    rs("Sal3").value = IIf(Trim(Me.TxtPrePayd(1).text) = "", Null, val(Me.TxtPrePayd(1).text))
    rs("Custom3").value = IIf(Trim(Me.TxtPrePayd(2).text) = "", Null, val(Me.TxtPrePayd(2).text))
    rs("TicketValue3").value = IIf(Trim(Me.TxtPrePayd(3).text) = "", Null, val(Me.TxtPrePayd(3).text))
    rs("AddOther3").value = IIf(Trim(Me.TxtPrePayd(4).text) = "", Null, val(Me.TxtPrePayd(4).text))
    rs("net3").value = IIf(Trim(Me.TxtPrePayd(5).text) = "", Null, val(Me.TxtPrePayd(5).text))
    rs("ValEndService3").value = IIf(Trim(Me.TxtPrePayd(6).text) = "", Null, val(Me.TxtPrePayd(6).text))
    rs("AdvanceTotal3").value = IIf(Trim(Me.TxtPrePayd(7).text) = "", Null, val(Me.TxtPrePayd(7).text))
    rs("VlueVaction3").value = IIf(Trim(Me.TxtPrePayd(8).text) = "", Null, val(Me.TxtPrePayd(8).text))
    rs("Cash3").value = IIf(Trim(Me.TxtPrePayd(9).text) = "", Null, val(Me.TxtPrePayd(9).text))
    rs("Discounts3").value = IIf(Trim(Me.TxtPrePayd(10).text) = "", Null, val(Me.TxtPrePayd(10).text))
    rs("TotalDis3").value = IIf(Trim(Me.TxtPrePayd(11).text) = "", Null, val(Me.TxtPrePayd(11).text))
    
    ''///
     rs("TotlPaidEndSer").value = IIf(Trim(Me.TxtTotlPaidEndSer.text) = "", Null, val(Me.TxtTotlPaidEndSer.text))
     rs("ManulaNO").value = IIf(Trim(Me.TxtManulaNO.text) = "", Null, Trim(Me.TxtManulaNO.text))
     rs("ManualNO").value = IIf(Trim(Me.TxtManulaNO.text) = "", Null, Trim(Me.TxtManulaNO.text))
     rs("Due").value = IIf(val(Me.TxtDue.text) = 0, Null, val(Me.TxtDue.text))
     rs("Discounts").value = IIf(val(Me.TxtPrePayd(13).text) = 0, Null, val(Me.TxtPrePayd(13).text))
     rs("Discounts2").value = IIf(val(Me.TxtPrePayd(12).text) = 0, Null, val(Me.TxtPrePayd(12).text))
     
     rs("DisSalary2").value = IIf(val(Me.TxtPrePayd(14).text) = 0, Null, val(Me.TxtPrePayd(14).text))
     rs("DisSalaryPayed").value = IIf(val(Me.TxtPrePayd(15).text) = 0, Null, val(Me.TxtPrePayd(15).text))
     rs("DisSalary").value = IIf(val(Me.TxtPrePayd(16).text) = 0, Null, val(Me.TxtPrePayd(16).text))
     If Option8(1).value = True Then
     rs("Fina_Pruch").value = 1
     Else
     rs("Fina_Pruch").value = 0
     End If
     If Option1.value = True Then
        rs("NCashingType").value = 1
      ElseIf Option3.value = True Then
        rs("NCashingType").value = 3
     Else
         rs("NCashingType").value = 0
     End If
                 If Option5.value = True Then
                rs("salary_or_advance").value = 1
               ' XPTxtID1.Text = 0
            If val(XPTxtID1.text) = 0 Then
                XPTxtID1.text = CStr(new_id("TblEmpAdvance", "AdvanceID", "", True))
            End If

            rs("AdvanceID").value = val(XPTxtID1.text)
       '     rs.update
        End If
    rs("TotalNotesValue").value = myRound(Me.txtTotalWithVat.text)
    rs("VAT").value = IIf(TxtVATValue.text = "", Null, myRound(TxtVATValue.text))
        '         rs("AdvanceID").value = Val(XPTxtID1.text)
        rs("Prefix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)
        rs("EmpIDD").value = EmpIDD
        If DCboCashType.ListIndex = 12 Then
        rs("VATVowalNo").value = IIf(Trim(Me.TxtEndService.text) = "", Null, Trim(Me.TxtEndService.text))
        Else
        rs("TxtEndService").value = IIf(Trim(Me.TxtEndService.text) = "", Null, Trim(Me.TxtEndService.text))
        End If
        rs("Order_no").value = IIf(Trim(Me.TXT_order_no.text) = "", Null, Trim(Me.TXT_order_no.text))
        rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
        rs("NumIqama").value = IIf(Trim(Me.TxtNumIqama.text) = "", Null, Trim(Me.TxtNumIqama.text))
        rs("Telephone").value = IIf(Trim(Me.TxtTelephone.text) = "", Null, Trim(Me.TxtTelephone.text))
        rs("OrderIDD").value = IIf(TxtOrder.text = "", Null, (TxtOrder.text))
        rs("AdvanceIDD").value = IIf(TxtAdvance.text = "", Null, val(TxtAdvance.text))
        rs("branch_no").value = val(Me.dcBranch.BoundText)
        rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, val(XPTxtVal.text))
        rs("note_value_by_characters").value = IIf(lbl(18).Caption = "", Null, lbl(18).Caption)
        ''////////////
        If DCboCashType.ListIndex = 7 Or DCboCashType.ListIndex = 1 Or DCboCashType.ListIndex = 5 Or DCboCashType.ListIndex = 2 Then
        rs("PreVAT").value = val(Me.TxtPrePayd(17).text)
        Else
        rs("PreVAT").value = 0
        End If
        If IncludVAT.value = vbChecked Then
        rs("IncludVAT").value = 1
        Else
        rs("IncludVAT").value = 0
        End If
        
        rs("TicktConract").value = val(Me.TxtTicktConract.text)
        
        If DCboCashType.ListIndex = 8 Or DCboCashType.ListIndex = 10 Then
        rs("Total2").value = val(Me.txtTotal2.text)
        rs("Sal2").value = val(Me.txtSal2.text)
        rs("net2").value = val(Me.txtNet2.text)
        rs("VlueVaction2").value = val(Me.TxtVlueVaction2.text)
        rs("TicketValue2").value = val(Me.txtTicketValue2.text)
        rs("Custom2").value = val(Me.txtCustom2.text)
        rs("AdvanceTotal2").value = val(Me.TXTAdvanceTotal2.text)
        rs("Cash2").value = val(Me.TxtCash2.text)
        rs("Salary2").value = val(Me.txtSalary2.text)
        rs("SalEntitOther2").value = val(Me.TxtSalEntitOther2.text)
        rs("Other2").value = val(Me.TxtOther2.text)
        rs("Advance12").value = val(Me.txtAdvance12.text)
        rs("InsuranceValue2").value = val(Me.TxtInsuranceValue2.text)
        rs("LastTotal2").value = val(Me.TXTLastTotal2.text)
        rs("SalaryVocation2").value = val(Me.txtSalaryVocation2.text)
        rs("ValueTickt2").value = val(Me.txtValueTickt2.text)
        rs("Totalsalary2").value = val(Me.TxtTotalsalary2.text)
        rs("CusTiket2").value = val(Me.TxtCusTiket2.text)
        rs("CusTiket").value = val(Me.TxtCusTiket.text)
        rs("AddOther").value = val(Me.TxtAddOther.text)
        rs("AddOther2").value = val(Me.TxtAddOther2.text)
        rs("ValEndService2").value = val(Me.TxtValEndService2.text)
        rs("ValEndService").value = val(Me.TxtValEndService.text)
        rs("TotalDis2").value = val(Me.TxtTotalDis2.text)
        rs("TotalDis").value = val(Me.TxtTotalDis.text)
        rs("ValEndService").value = val(Me.TxtValEndService.text)
        
        rs("AdvanceTotal").value = val(Me.TXTAdvanceTotal.text)
        rs("VlueVaction").value = val(Me.TxtVlueVaction.text)
        rs("Cash").value = val(Me.TxtCash.text)
        rs("LastTotal2").value = val(Me.TXTLastTotal2.text)
        rs("LastTotal").value = val(Me.TXTLastTotal.text)
        rs("CusTiket").value = val(Me.TxtCusTiket.text)
        rs("Custom").value = val(Me.txtCustom.text)
        rs("TicketValue").value = val(Me.txtTicketValue.text)
        rs("Sal").value = val(Me.txtSal.text)
        rs("total").value = val(Me.txtTotal.text)
        rs("EmpIDEnd").value = val(Me.DcbEmpEndService.BoundText)
        rs("BrnchIDEnd").value = val(Me.DcbBranchEndServ.BoundText)
        rs("EmpIDVac").value = val(Me.DcboEmpName.BoundText)
        rs("BrnchIDVac").value = val(Me.dcBranch1.BoundText)
        rs("Salary33").value = val(Me.TxtSalary.text)
        rs("SalEntitOther").value = val(Me.TxtSalEntitOther.text)
        rs("Other").value = val(Me.Txtother.text)
        rs("Advance1").value = val(Me.txtAdvance1.text)
        rs("InsuranceValue").value = val(Me.TxtInsuranceValue.text)
       rs("SalaryVocation").value = val(Me.txtSalaryVocation.text)
       rs("ValueTickt").value = val(Me.txtValueTickt.text)
        End If
      If DCboCashType.ListIndex = 8 Then
      rs("Salary3").value = val(Me.txtSalary3.text)
      rs("SalEntitOther3").value = val(Me.TxtSalEntitOther3.text)
      rs("SalaryVocation3").value = val(Me.txtSalaryVocation3.text)
      rs("Other3").value = val(Me.TxtOther3.text)
      rs("ValueTickt3").value = val(Me.txtValueTickt3.text)
      rs("Advance13").value = val(Me.txtAdvance13.text)
      rs("InsuranceValue3").value = val(Me.TxtInsuranceValue3.text)
      End If
      
    '**********************************12 12 2015
    rs("PayrollMonth").value = Me.CmbMonth1.ListIndex + 1
    rs("PayrollYear").value = val(Me.CboYear1.ListIndex)
    rs("empDes1").value = IIf(empDes1.text = "", "", Trim(empDes1.text))
    rs("empDes").value = IIf(empDes.text = "", "", Trim(empDes.text))
    rs("PayDes").value = IIf(PayDes.text = "", "", Trim(PayDes.text))
    rs("TxtNoSupplerDes").value = IIf(TxtNoSupplerDes.text = "", "", Trim(TxtNoSupplerDes.text))
    rs("TxtOrderSuppler").value = IIf(TxtOrderSuppler.text = "", Null, val((TxtOrderSuppler.text)))
    ''''//////////////
    rs("Month_Req").value = IIf(val(dcMontth.BoundText) = 0, Null, Trim(dcMontth.BoundText))
    rs("Year_Req").value = IIf(val(dcDur.BoundText) = 0, Null, Trim(dcDur.BoundText))
    rs("Brnch_Req").value = IIf(val(DcbBrReq.BoundText) = 0, Null, Trim(DcbBrReq.BoundText))
    
     rs("BeneficiaryBanck").value = IIf(TxtBeneficiaryBanck.text = "", "", Trim(TxtBeneficiaryBanck.text))
     rs("BeneficiaryAddress").value = IIf(TxtBeneficiaryAddress.text = "", "", Trim(TxtBeneficiaryAddress.text))
     rs("BeneficiaryACNo").value = IIf(TxtBeneficiaryACNo.text = "", "", Trim(TxtBeneficiaryACNo.text))
     rs("BenefiBanckAddress").value = IIf(TxtBenefiBanckAddress.text = "", "", Trim(TxtBenefiBanckAddress.text))
     rs("BenefiBanckCode").value = IIf(TxtBenefiBanckCode.text = "", "", Trim(TxtBenefiBanckCode.text))
     rs("RemitterName").value = IIf(TxtRemitterName.text = "", "", Trim(TxtRemitterName.text))
     rs("ReportName").value = IIf(TxtReportName.text = "", "", Trim(TxtReportName.text))
    '''///////////
    rs("BenefTelephone").value = IIf(TxtBenefTelephone.text = "", Null, (TxtBenefTelephone.text))
    '''//////
    rs("InsuranceValue").value = IIf(Me.TxtInsuranceValue.text = "", 0, val(TxtInsuranceValue.text))
    rs("BenefPlaceBrith").value = IIf(TxtBenefPlaceBrith.text = "", Null, (TxtBenefPlaceBrith.text))
    rs("BenefBrithDate").value = BenefBrithDate.value
    rs("BenefDateExpEqama").value = BenefDateExpEqama.value
    rs("KafelAddress").value = IIf(TxtKafelAddress.text = "", Null, (TxtKafelAddress.text))
    rs("BenefPlaceIqama").value = IIf(TxtBenefPlaceIqama.text = "", Null, (TxtBenefPlaceIqama.text))
    rs("BenefIBAN").value = IIf(TxtBenefIBAN.text = "", Null, (TxtBenefIBAN.text))
    ''''//
    
     rs("BenefNumIqama").value = IIf(TxtBenefNumIqama.text = "", "", Trim(TxtBenefNumIqama.text))
     rs("BenefCountry").value = IIf(TXtBenefCountry.text = "", "", Trim(TXtBenefCountry.text))
     rs("BenefCity").value = IIf(TxtBenefCity.text = "", "", Trim(TxtBenefCity.text))
     rs("BenefStreet").value = IIf(TxtBenefStreet.text = "", "", Trim(TxtBenefStreet.text))
     rs("BenefGovernorate").value = IIf(TxtBenefGovernorate.text = "", "", Trim(TxtBenefGovernorate.text))
     rs("Country").value = IIf(TxtCountry.text = "", "", Trim(TxtCountry.text))
     rs("Governorate").value = IIf(TxtGovernorate.text = "", "", Trim(TxtGovernorate.text))
     rs("City").value = IIf(TxtCity.text = "", "", Trim(TxtCity.text))
     rs("Street").value = IIf(TxtStreet.text = "", "", Trim(TxtStreet.text))
     rs("Adress2").value = IIf(TxtAdress2.text = "", "", Trim(TxtAdress2.text))
     rs("KafelName").value = IIf(TxtKafelName.text = "", "", Trim(TxtKafelName.text))
     rs("KafeltEL").value = IIf(TxtKafeltEL.text = "", "", Trim(TxtKafeltEL.text))
     
     rs("TradingContractID").value = IIf(txtTradingContractID.text = "", 0, val(txtTradingContractID.text))
               
       
     '**********************************12 12 2015
     
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        rs("general_des_notes").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    
        rs("person").value = IIf(Me.txtperson.text = "", "", Me.txtperson.text)

        rs("NoteType").value = 5
        rs("TransferExpenses").value = val(txtTransferExpenses.text)
        'TransferExpenses
    
      rs("DeptID").value = IIf(Me.DcbDepartment.text = "", Null, DcbDepartment.BoundText)
          rs("NoteDate").value = XPDtbTrans.value
       ' rs("NoteDate").value = Format$(Date, "dd-mm-yyyy")
        rs("NoteDateH").value = Me.Txt_DateHigri.value
   
        rs("CashingType").value = IIf(DCboCashType.ListIndex = -1, Null, DCboCashType.ListIndex)

        If DCboCashType.ListIndex = 3 Then
            rs("ProjectMainID").value = IIf(val(DBCboClientName.BoundText) = 0, Null, val(DBCboClientName.BoundText))

        ElseIf DCboCashType.ListIndex = 4 Then
            rs("EmpAccountCode").value = IIf(DBCboClientName.text = "", Null, DBCboClientName.BoundText)
            rs("CusID").value = Null

            rs("person").value = IIf(DBCboClientName.text = "", "", Trim(DBCboClientName.text))
        
        ElseIf DCboCashType.ListIndex = 5 Or DCboCashType.ListIndex = 6 Then
            rs("BTCashAccountcode").value = IIf(DBCboClientName.text = "", Null, DBCboClientName.BoundText)
            rs("CusID").value = Null
        Else
            rs("CusID").value = IIf(DBCboClientName.text = "", Null, DBCboClientName.BoundText)
 
        End If
    
        If Option4.value = True Then
            rs("salary_or_advance").value = 0
    
        ElseIf Option6.value = True Then
            rs("salary_or_advance").value = 2
        ElseIf Option7.value = True Then
            rs("salary_or_advance").value = 3
              
        Else
            rs("salary_or_advance").value = Null
        End If
    If val(DCboCashType.ListIndex) = 10 Then
    sql = "Update End_of_service set PaymPaid=1 where id=" & val(Me.TxtEndService.text) & " "
    Cn.Execute sql
    End If
        'DcboBox
        If Me.ChkTrans.value = vbChecked Then
            If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                rs("Transaction_ID").value = val(Me.TxtTransID.text)
            End If

        Else
            rs("Transaction_ID").value = Null
        End If

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
       
        ElseIf Me.CboPayMentType.ListIndex = 2 Then
            rs("NoteCashingType").value = 2
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("ChequeBoxID").value = Null
  
        ElseIf Me.CboPayMentType.ListIndex = 3 Then
            rs("NoteCashingType").value = 3
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("ChequeBoxID").value = Null
      ElseIf Me.CboPayMentType.ListIndex = 4 Then
           rs("NoteCashingType").value = 4
      ElseIf Me.CboPayMentType.ListIndex = 5 Then
           rs("NoteCashingType").value = 5
        End If
''//
       rs("ProjectID").value = val(dcproject1.BoundText)
       rs("Pand").value = val(Dcterm1.BoundText)
       rs("Oper").value = val(Me.dcopr.BoundText)
''/

        rs("UserID").value = user_id
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        rs("foxy_no").value = val(Text1.text)
        rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        rs("numbering_type1").value = sand_numbering_type(4) '”‰œ «·œ›⁄
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)
        rs("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
    
  '**********
    If subContOpt(0).value = True Then
      rs("subcontractType").value = 0
    ElseIf subContOpt(1).value = True Then
    rs("subcontractType").value = 1
  ElseIf subContOpt(2).value = True Then
    rs("subcontractType").value = 2
    Else
    rs("subcontractType").value = 2
    End If
    '*********
     saveBillVendor
     If Not saveBillBuy Then GoTo ErrTrap
     saveBillProject
      If val(DCboCashType.ListIndex) = 12 Then
  Cn.Execute "update TblVATAvowal set Paid =1 where ID=" & val(TxtEndService.text) & ""
  End If


        rs.update
        
        If SystemOptions.IsCheque = True And CboPayMentType.ListIndex = 1 Then GoTo endSave
 
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            Set RsDev = New ADODB.Recordset
        '    RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                            StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
            Line1 = setfoxy_Line
            Line2 = setfoxy_Line
            Line3 = setfoxy_Line
Dim SngTemp2 As Double
Dim SngTemp As Double
Dim SngTemp3 As Double

            '«·ÿ—› «·„œÌ‰
            ' ›Ì Õ«·… «·ÕÊ«·«  «·»‰ﬂÌ… ÊÊÃÊœ „’—Ê›«  »‰ﬂÌ… ⁄·Ì⁄«
            If (CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 4 Or CboPayMentType.ListIndex = 5) And val(Me.txtTransferExpenses.text) > 0 Then
                RsDev.AddNew
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = 10000
                RsDev("DEV_ID_Line_No1").value = Line2
                
                If CboPayMentType.ListIndex = 2 Then
                RsDev("Account_Code").value = Account_Code_dynamic
                Else
                RsDev("Account_Code").value = DcboCreditSide.BoundText
                End If
                RsDev("NextAccount_Code").value = DcboDebitSide.BoundText
              
                'If val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 1 Then
                If IncludVAT.value = vbChecked Then
                GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
                PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
               
               ' Percetage = 0
                If Percetage = 0 Then
                    Percetage = 1
                Else
                    Percetage = Percetage / 100 + 1
                End If
                
                SngTemp3 = val(Me.txtTransferExpenses.text) / Percetage
                Else
                SngTemp3 = val(Me.txtTransferExpenses.text)
                End If
                SngTemp3 = val(Format(val(SngTemp3), "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
                RsDev("Value").value = SngTemp3
                RsDev("valuee").value = (SngTemp3) / val(TxtCurrencyRate.text)
                RsDev("rate").value = val(Me.TxtCurrencyRate.text)
                RsDev("currency").value = DcbCurrency.text
               ' Else
               ' RsDev("Value").value = val(Me.txtTransferExpenses.Text)
               ' End If
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = "  „’«—Ì› ÕÊ«·Â »‰ﬂÌ…  " & XPMTxtRemarks.text & CHR(13) & txt_general_des
                RsDev("Notes_ID").value = val(XPTxtID.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
             
                RsDev.update
                ''///////////
                If IncludVAT.value = vbChecked Then
                    GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
                    PercentgValueAddedAccount_Transec XPDtbTrans.value, 23, 0, AccountVATCreit, Percetage
                  
                    If Percetage = 0 Then GoTo 235
                
                   RsDev.AddNew
                         Line1 = setfoxy_Line
            Line2 = setfoxy_Line
            Line3 = setfoxy_Line
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = 10001
                RsDev("DEV_ID_Line_No1").value = Line2
                
               
                
                
                
                SngTemp3 = SngTemp3 * Percetage / 100
                SngTemp3 = val(Format(val(SngTemp3), "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
                If val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 1 Then
                RsDev("Value").value = SngTemp3
                RsDev("valuee").value = SngTemp3 / val(TxtCurrencyRate.text)
                RsDev("rate").value = val(Me.TxtCurrencyRate.text)
                RsDev("currency").value = DcbCurrency.text
                Else
                RsDev("Value").value = SngTemp3
                End If
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = " Õ”«» «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì…. „’«—Ì› ÕÊ«·Â »‰ﬂÌ…   " & XPMTxtRemarks.text & CHR(13) & txt_general_des
                RsDev("Notes_ID").value = val(XPTxtID.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("Account_Code").value = AccountVATCreit
                RsDev("NextAccount_Code").value = DcboCreditSide.BoundText
                RsDev.update
                End If
        
            End If
235:
      total_value = val(txtTransferExpenses.text) + val(txtVat2.text)
      total_valuee = total_value
      
     If total_value > 0 Then
        Msg = "„’«—Ì› «·ﬁÌ„… «·„÷«›… "
     If ModAccounts.AddNewDev(LngDevID, 10002, DcboCreditSide.BoundText, total_value, 1, Msg, val(XPTxtID.text), , , , XPDtbTrans.value, user_id, , , , total_valuee, DcbCurrency.text, TxtCurrencyRate.text, , , Line2, , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                                                   
    End If
    End If
''//////////////
            If val(TxtPrePayd(17).text) > 0 And (val(DCboCashType.ListIndex) = 1 Or val(DCboCashType.ListIndex) = 2 Or val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 5 Or val(DCboCashType.ListIndex) = 7 Or Option3.value = True) Then
                
                   RsDev.AddNew
                         Line1 = setfoxy_Line
            Line2 = setfoxy_Line
            Line3 = setfoxy_Line
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = 1003
                RsDev("DEV_ID_Line_No1").value = Line2
                GetValueAddedAccount XPDtbTrans.value, AccountVATCreit, , 1, 23
                SngTemp3 = val(TxtPrePayd(17).text)
                SngTemp3 = val(Format(SngTemp3, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
                If (val(DCboCashType.ListIndex) = 1 Or val(DCboCashType.ListIndex) = 2 Or val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 5 Or val(DCboCashType.ListIndex) = 7) And Option3 Then
                
                RsDev("Value").value = SngTemp3
                RsDev("valuee").value = SngTemp3 / val(TxtCurrencyRate.text)
                RsDev("rate").value = val(Me.TxtCurrencyRate.text)
                RsDev("currency").value = DcbCurrency.text
                Else
                RsDev("Value").value = SngTemp3
                End If
                RsDev("Credit_Or_Debit").value = 0
                If SystemOptions.UserInterface = ArabicInterface Then
                RsDev("Double_Entry_Vouchers_Description").value = " Õ”«» «·ﬁÌ„… «·„÷«›… ··„⁄«„·«  «·„«·Ì….   „œ›Ê⁄«  „ﬁœ„…   " & XPMTxtRemarks.text & CHR(13) & txt_general_des
                Else
                RsDev("Double_Entry_Vouchers_Description").value = "Vat  " & XPMTxtRemarks.text & CHR(13) & txt_general_des
                End If
                RsDev("Notes_ID").value = val(XPTxtID.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                RsDev("Account_Code").value = AccountVATCreit
                RsDev("NextAccount_Code").value = DcboCreditSide.BoundText
                RsDev.update
                End If
        
           
            '44444444444444444444444444
            Dim LastLine As Integer
If DCboCashType.ListIndex = 12 Then
LastLine = payGlVAT(LngDevID, val(XPTxtID.text))
GoTo ll
End If

If DCboCashType.ListIndex = 6 Then
LastLine = payGl(LngDevID, val(XPTxtID.text))
GoTo ll
End If
If DCboCashType.ListIndex = 7 Then
LastLine = Me.payGl122(LngDevID, val(XPTxtID.text))
GoTo ll
End If
If DCboCashType.ListIndex = 9 Then
LastLine = Me.payGl1Suppler(LngDevID, val(XPTxtID.text))
GoTo ll
End If

If DCboCashType.ListIndex = 8 Then
LastLine = Me.payGl8(LngDevID, val(XPTxtID.text))
If LastLine = -1 Then MsgBox "·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·ÊÃÊœ ﬁÌ„ ”«·»…": GoTo ErrTrap

GoTo ll
End If
If DCboCashType.ListIndex = 10 Then
LastLine = Me.payGl10(LngDevID, val(XPTxtID.text))
GoTo ll
End If
If DCboCashType.ListIndex = 11 Then
LastLine = Me.payGl16(LngDevID, val(XPTxtID.text))
GoTo ll
End If

 If DCboCashType.ListIndex = 1 And val(Label16.Caption) > 0 Then
LastLine = payGl1(LngDevID, val(XPTxtID.text))
GoTo llx
End If
  If DCboCashType.ListIndex = 1 And val(Label27(1).Caption) > 0 Then
LastLine = payGlBillBuy1(LngDevID, val(XPTxtID.text))
GoTo llx
End If
      
        Dim BranchID As Integer
         Dim BranchID2 As Integer
         Dim DeptSide As String
         Dim credit_side As String
         
                  BranchID = val(Me.dcBranch.BoundText)
            If DCboCashType.ListIndex = 4 Then
                BranchID2 = val(Me.DcbEmpBranch.BoundText)
             ElseIf DCboCashType.ListIndex = 1 Or DCboCashType.ListIndex = 0 Then
             BranchID2 = GetBrnchCustomer(val(DBCboClientName.BoundText))
             End If
             

                                  DeptSide = getBranchCurrentAccount(BranchID)
                                                 credit_side = getBranchCurrentAccount(BranchID2)
                                                 

            RsDev.AddNew
   If DCboCashType.ListIndex = 4 Or DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1 Then
                    If BranchID <> BranchID2 And BranchID2 <> 0 And DCboCashType.ListIndex = 4 Then
                        RsDev("branch_id").value = val(Me.DcbEmpBranch.BoundText)
                    ElseIf BranchID <> BranchID2 And (DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1) Then
                        RsDev("branch_id").value = GetBrnchCustomer(val(DBCboClientName.BoundText))
                    Else
                       RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                    End If
   
   Else
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
   End If
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 1
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
            RsDev("NextAccount_Code").value = DcboCreditSide.BoundText
            If val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 1 Then
                RsDev("Value").value = val(Me.XPTxtVal.text)
                RsDev("valuee").value = val(Me.XPTxtVal.text) / val(TxtCurrencyRate.text)
                RsDev("rate").value = val(Me.TxtCurrencyRate.text)
                RsDev("currency").value = DcbCurrency.text
            Else
                 RsDev("Value").value = val(Me.XPTxtVal.text)
            End If
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & txt_general_des ' XPMTxtRemarks.text
               RsDev("Departementid").value = IIf(Me.DcbDepartment.text = "", Null, DcbDepartment.BoundText)
            If DCboCashType.ListIndex = 3 Then
                Dim project_id As Integer
                'project_id = get_project_id(DBCboClientName.BoundText, "expanses_account")
                RsDev("projectid").value = val(DBCboClientName.BoundText)
                RsDev("Double_Entry_Vouchers_Description").value = "’—› ⁄·Ï „‘—Ê⁄" & DBCboClientName.text
           Else
                RsDev("projectid").value = val(dcproject1.BoundText)
            End If
            
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID

            RsDev.update
            


            
ll:
            '«·ÿ—› «·œ«∆‰
            LngDevID = LngDevID + 1
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
            RsDev("NextAccount_Code").value = DcboDebitSide.BoundText
            If (val(DCboCashType.ListIndex) = 1 Or val(DCboCashType.ListIndex) = 2 Or val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 5 Or val(DCboCashType.ListIndex) = 7) And Option3 Then
              '  RsDev("Value").value = (val(Me.XPTxtVal.Text) + val(Me.txtTransferExpenses.Text))
              If val(DCboCashType.ListIndex) = 1 Or val(DCboCashType.ListIndex) = 2 Or val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 5 Or val(DCboCashType.ListIndex) = 7 And Option3 Then
               ' RsDev("Value").value = (val(Me.XPTxtVal.Text) + val(TxtPrePayd(17).Text))
               '  „ «÷«›… «·ﬁÌ„… «·„÷«›… ›Ï «·ÿ—› «·«ŒÌ— „‰ «·ﬁÌœ
                RsDev("Value").value = (val(Me.XPTxtVal.text))
                '  „ «÷«›… «·ﬁÌ„… «·„÷«›… ›Ï «·ÿ—› «·«ŒÌ— „‰ «·ﬁÌœ
                'RsDev("valuee").value = (val(Me.XPTxtVal.Text) + val(TxtPrePayd(17).Text)) / val(TxtCurrencyRate.Text)
                RsDev("valuee").value = (val(Me.XPTxtVal.text)) / val(TxtCurrencyRate.text)
               Else
               RsDev("Value").value = (val(Me.XPTxtVal.text))
                RsDev("valuee").value = (val(Me.XPTxtVal.text)) / val(TxtCurrencyRate.text)
            End If
                RsDev("rate").value = val(Me.TxtCurrencyRate.text)
                RsDev("currency").value = DcbCurrency.text
            Else
                If val(DCboCashType.ListIndex) = 7 Then
           ' RsDev("Value").value = (val(Me.XPTxtVal.Text) + val(Me.txtTransferExpenses.Text))
                    RsDev("Value").value = val(Me.XPTxtVal.text) + val(Format(val(TxtPrePayd(17)), "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
                    
                Else
                    RsDev("Value").value = val(Me.XPTxtVal.text)
                End If
            End If
            RsDev("Credit_Or_Debit").value = 1
            
            If (CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 4 Or CboPayMentType.ListIndex = 5) And val(Me.txtTransferExpenses.text) > 0 Then
                RsDev("DEV_ID_Line_No").value = 3
                RsDev("DEV_ID_Line_No1").value = Line3
            Else
                RsDev("DEV_ID_Line_No").value = 2
                RsDev("DEV_ID_Line_No1").value = Line2
            End If

   If DCboCashType.ListIndex = 6 Then
        RsDev("DEV_ID_Line_No").value = LastLine + 1
                RsDev("DEV_ID_Line_No1").value = LastLine + 1
   End If
   
      If DCboCashType.ListIndex = 1 And LastLine > 0 Then
        RsDev("DEV_ID_Line_No").value = LastLine + 1
                RsDev("DEV_ID_Line_No1").value = LastLine + 1
   End If
   
   
      If DCboCashType.ListIndex = 7 Then
        RsDev("DEV_ID_Line_No").value = LastLine + 1
                RsDev("DEV_ID_Line_No1").value = LastLine + 1
   End If
      If DCboCashType.ListIndex = 11 Then
        RsDev("DEV_ID_Line_No").value = LastLine + 1
        RsDev("DEV_ID_Line_No1").value = LastLine + 1
   End If
   'If val(Label16.Caption) > 0 Then
   '   RsDev("DEV_ID_Line_No").value = LastLine + 1
   '             RsDev("DEV_ID_Line_No1").value = LastLine + 1
   'End If
   
      If DCboCashType.ListIndex = 8 Then
        RsDev("DEV_ID_Line_No").value = LastLine + 1
                RsDev("DEV_ID_Line_No1").value = LastLine + 1
   End If
   If DCboCashType.ListIndex = 9 Then
    RsDev("DEV_ID_Line_No").value = LastLine + 1
                RsDev("DEV_ID_Line_No1").value = LastLine + 1
End If
   If DCboCashType.ListIndex = 10 Then
    RsDev("DEV_ID_Line_No").value = LastLine + 1
                RsDev("DEV_ID_Line_No1").value = LastLine + 1
End If
   If DCboCashType.ListIndex = 12 Then
    RsDev("DEV_ID_Line_No").value = LastLine + 1
                RsDev("DEV_ID_Line_No1").value = LastLine + 1
End If

   
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & txt_general_des ' XPMTxtRemarks.text
            ' RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
       If DCboCashType.ListIndex = 4 Or DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1 Then
 

                                       '          DeptSide1 = DcboDebitSide.BoundText
                                       '          CreditSide1 = DcboCreditSide.BoundText
                                                 Msg = XPMTxtRemarks.text & CHR(13) & txt_general_des ' XPMTxtRemarks.text
       If BranchID <> BranchID2 And BranchID2 <> 0 And SystemOptions.DontDistributeLegalACC = False Then
 total_value = val(Me.XPTxtVal.text) ' + val(Me.txtTransferExpenses.Text)
LastLine = 4
OtherInformation.NextAccount_Code = DeptSide
                                               If ModAccounts.AddNewDev(LngDevID, LastLine, credit_side, total_value, 0, Msg, val(XPTxtID.text), , , , XPDtbTrans.value, user_id, , , , , , , , , CDbl(LastLine), , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              'LastLine = LastLine + 1
                                                              LastLine = 5
                                                              OtherInformation.NextAccount_Code = credit_side
                                                        '????
                                                              If ModAccounts.AddNewDev(LngDevID, LastLine, DeptSide, total_value, 1, Msg, val(XPTxtID.text), , , , XPDtbTrans.value, user_id, , , , , , , , , CDbl(LastLine), , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              
                                                        
                                    
                                                        
                                      LastLine = LastLine + 1
        

       End If
       End If
       ''//////////////////////
      If val(TxtPrePayd(17).text) > 0 Then
              If DCboCashType.ListIndex = 4 Or DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1 Then
                                                 
                                                If SystemOptions.UserInterface = ArabicInterface Then
                                                 Msg = XPMTxtRemarks.text & CHR(13) & txt_general_des & "ﬁÌ„… „÷«›…"  ' XPMTxtRemarks.text
                                                 Else
                                    Msg = XPMTxtRemarks.text & CHR(13) & txt_general_des & " VAT  "  ' XPMTxtRemarks.text
                                                 End If
                                                 
       If BranchID <> BranchID2 Then
 total_value = val(Format(val(val(TxtPrePayd(17).text)), "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
  
LastLine = 6
OtherInformation.NextAccount_Code = DeptSide
                                               If ModAccounts.AddNewDev(LngDevID, LastLine, credit_side, total_value, 0, Msg, val(XPTxtID.text), , , , XPDtbTrans.value, user_id, , , , , , , , , CDbl(LastLine), , , , , , , , , BranchID, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              'LastLine = LastLine + 1
                                                              LastLine = 7
                                                              OtherInformation.NextAccount_Code = credit_side
                                                        '????
                                                              If ModAccounts.AddNewDev(LngDevID, LastLine, DeptSide, total_value, 1, Msg, val(XPTxtID.text), , , , XPDtbTrans.value, user_id, , , , , , , , , CDbl(LastLine), , , , , , , , , BranchID2, , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                                                                   
                                                              End If
                                                              
                                                        
                                    
                                                        
                                      LastLine = LastLine + 1
        

       End If
       End If
  End If
  

 ''  /////////////////
        
            LblDevID.Caption = LngDevID
            lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
        End If
endSave:
          save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.text, "„œ›Ê⁄« ", Me.XPDtbTrans.value
        save_cost_center
        ' If DCboCashType.ListIndex = 7 Or DCboCashType.ListIndex = 1 Then
         If (val(DCboCashType.ListIndex) = 1 Or val(DCboCashType.ListIndex) = 2 Or val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 5 Or val(DCboCashType.ListIndex) = 7) And Option3 Then
           updateNotesValueAndNobytext val(XPTxtID.text), Format(XPTxtVal.text, "###.00") + val(txtVat2.text)
           Else
           updateNotesValueAndNobytext val(XPTxtID.text), Format(XPTxtVal.text, "###.00")
       End If
      
        '     End If
        
        'Õ›Ÿ «·„’«—Ì› › ÃœÊ· «·„œ›Ê⁄«  Ê «·„ﬁ»Ê÷« 
     
        If SavePaymentAndReciveDetails(0, TxtNoteSerial.text, TxtNoteSerial1.text, TXT_order_no.text, XPDtbTrans.value) = True Then
        End If

        'Õ›Ÿ »Ì«‰«  «·”·›…
        
           
'    StrSQL = "Delete From TblEmpAdvance Where AdvanceID=" & val(Me.XPTxtID1.Text)
'                Cn.Execute StrSQL, , adExecuteNoRecords
         
            StrSQL = "Delete From TblEmpAdvanceDetails Where AdvanceID=" & val(Me.XPTxtID1.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
             
     

        saveAdvancedData
        
    End If
updateNotesValueAndNobytext val(XPTxtID.text), Format(XPTxtVal.text, "###.00")

   

    If Option1.value = True Then
    
        FIFO_FUNCTION val(DBCboClientName.BoundText)
    End If
   
    If Option2.value Then
        Distribute_to_bills Me.lblsqlstring, val(DBCboClientName.BoundText)
    End If
 
  If val(DCboCashType.ListIndex) = 8 Then
  Cn.Execute "update TblVocationEntitlements set PayedPayment =1 where ID=" & val(TxtDue.text) & " and not (NoteSerial is null)"
  End If
    If val(DCboCashType.ListIndex) = 4 And Option5.value = True And val(TxtAdvance.text) > 0 Then
  Cn.Execute "update TblEmpAdvanceRequest set AccAproved=1  where AdvanceID=" & val(TxtAdvance.text) & ""
  End If
  If val(DCboCashType.ListIndex) = 6 Then
  SaveSalaryPyment
  End If

llx:

        saveChequeBoxContents1 (val(XPTxtID.text))
        
   '     If val(Label16.Caption) > 0 Then
'LastLine = Me.payGl1(LngDevID, val(XPTxtID.text))
'GoTo ll
'End If


       Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
              rs.Resync adAffectCurrent
        CuurentLogdata
           XPTxtValE.text = Format(XPTxtValE.text, "#,##0.00")
          XPTxtVal.text = Format(XPTxtVal.text, "#,##0.00")
        Select Case Me.TxtModFlg.text

            Case "N"
 
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Operation data was saved " & CHR(13)
                    Msg = Msg + "need another operation"
        
                Else
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
 
                End If
          
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    MsgBox "Update success", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If

                lbl(45).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
              
        End Select

        TxtModFlg.text = "R"
        '«· Ê“Ì⁄ ⁄·Ï „—ﬂ“ «· ﬂ·›… «·⁄«„
        '     If Me.DcCostCenter.BoundText <> "" Then
      
  
   WriteCustomerBalPublic Me.DcboDebitSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
RetriveBillVendorData
RetriveBillBuyData
RetriveBillProjectData
    WriteInfo
    
'rs.Resync adAffectCurrent
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
        Msg = "Can Not Save Plaese Make sure of the validity of the data"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry an error occurred while saving"
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Function saveAdvancedData()

 Dim StrSQL As String
    If DCboCashType.ListIndex <> 4 Or Option5.value = False Then Exit Function
    
    Dim RsDetails As ADODB.Recordset
    Dim rs  As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'rs.Open "TblEmpAdvance", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     * from dbo.TblEmpAdvance Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    rs.AddNew
    rs("AdvanceID").value = val(XPTxtID1.text)
    rs("AdvanceDate").value = XPDtbTrans.value
    rs("Emp_ID").value = get_EMPLOYEEIdFromAccountCode(Me.DBCboClientName.BoundText, "Emp_ID")
    rs("AdvanceValue").value = IIf(XPTxtVal.text = "", Null, val(XPTxtVal.text))
    '    rs("BoxID").value = IIf(Me.DcboBox.BoundText = "", Null, Me.DcboBox.BoundText)
    rs("PaymentCounts").value = val(Me.TxtPaymentCounts.text)
    rs("AutoDiscount").value = IIf(Me.ChkSaleryDis.value = vbChecked, 1, 0)
    rs("FirstMonthPayment").value = Me.CmbMonth.ListIndex + 1
    rs("FirstYearPayment").value = val(Me.CboYear.text)
    rs("UserID").value = Me.DCboUserName.BoundText
    rs("AdvanceType").value = 0
    rs("RetrunID").value = Null
    rs.update

    Dim i  As Integer
    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblEmpAdvanceDetails Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    For i = Me.FG.FixedRows To FG.rows - 1

        If val(FG.TextMatrix(i, FG.ColIndex("PartNO"))) <> 0 Then
            RsDetails.AddNew
            RsDetails("AdvanceID").value = val(XPTxtID1.text)
            RsDetails("PartNO").value = FG.TextMatrix(i, FG.ColIndex("PartNO"))
            RsDetails("PartValue").value = FG.TextMatrix(i, FG.ColIndex("PartValue"))
            RsDetails("PartDate").value = FG.TextMatrix(i, FG.ColIndex("PartDate"))
            RsDetails.update
        End If

    Next i

End Function





Private Function SqlQ(ByVal s As String) As String
    SqlQ = Replace(s, "'", "''")
End Function

Private Function SqlDateOrNull(ByVal v As String) As String
    If Trim$(v) = "" Then
        SqlDateOrNull = "NULL"
    Else
        ' ·Ê «· «—ÌŒ ⁄‰œﬂ √’·« »Ì Œ“‰ ‰’ “Ì 20/01/2026 √Ê 2026-01-20
        ' Œ·ÌÂ Ì ÕÊ· ›Ì SQL »‘ﬂ· ¬„‰:
        SqlDateOrNull = "CONVERT(datetime,'" & SqlQ(v) & "',103)"
    End If
End Function

Private Function SqlNum(ByVal d As Double) As String
    ' ⁄‘«‰ «·›«’·… «·⁄‘—Ì… „«  ﬂ”—‘ SQL Õ”» ≈⁄œ«œ«  «·ÊÌ‰œÊ“
    SqlNum = Replace(CStr(d), ",", ".")
End Function

Public Function saveBillBuy() As Boolean
    On Error GoTo EH

    Dim StrSQL As String
    Dim i As Integer
    Dim Diff As Double
    Dim Remaining As Double
    Dim PayAmount As Double

    saveBillBuy = False

    '========================================================
    ' Transaction («Œ Ì«—Ì ·ﬂ‰ √›÷·)
    '========================================================
'    Cn.BeginTrans

    '========================================================
    ' 1) ·Ê  ⁄œÌ·: «„”Õ «·Õ—ﬂ«  «·ﬁœÌ„…
    '========================================================
    If Me.TxtModFlg.text = "E" Then
        StrSQL = "DELETE FROM TblNotesBillBuyPayment WHERE NoteID1=" & val(Me.XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords

        StrSQL = "DELETE FROM TblBillBuyPayment WHERE TypTrans IS NULL AND NoteID=" & val(Me.XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    '========================================================
    ' 2) ﬁÌ„… «·”œ«œ «·≈Ã„«·Ì…
    '========================================================
    PayAmount = val(Me.XPTxtVal.text)

    '========================================================
    ' 3) Ê“¯⁄ «·”œ«œ ⁄·Ï «·›Ê« Ì— «·„Œ «—… ›ﬁÿ
    '    Ê”Ã· ”ÿ— ·ﬂ· ›« Ê—… ›Ì TblNotesBillBuyPayment (INSERT „»«‘—)
    '========================================================
    With VSFlexGrid1

        For i = .FixedRows To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then

                         Remaining = val(.TextMatrix(i, .ColIndex("RemainingValue")))
                    If Remaining <= 0 Then GoTo ContinueLoop
                    
                    Diff = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
                    If Diff <= 0 Then GoTo ContinueLoop
                    
                    If Diff > Remaining Then
                        Diff = Remaining
                        .TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
                    End If
                    
                    .TextMatrix(i, .ColIndex("NetValue")) = Remaining - Diff



                ' «ﬂ » ⁄·Ï «·Ã—Ìœ
'                .TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
'                .TextMatrix(i, .ColIndex("NetValue")) = Remaining - Diff

                '================================================
                ' INSERT „»«‘— ›Ì TblNotesBillBuyPayment
                '================================================
                StrSQL = ""
                StrSQL = StrSQL & "INSERT INTO TblNotesBillBuyPayment "
                StrSQL = StrSQL & "(NoteID1, NoteID, branch_no, NoteSerial1, Note_Value, PayedValue, too, DueDate, NoteDate, TransPayedValue, NetValue, RemainingValue) "
                StrSQL = StrSQL & "VALUES ("
                StrSQL = StrSQL & val(Me.XPTxtID.text) & ","
                StrSQL = StrSQL & val(.TextMatrix(i, .ColIndex("NoteID"))) & ","
                StrSQL = StrSQL & val(.TextMatrix(i, .ColIndex("branch_no"))) & ","
                StrSQL = StrSQL & val(.TextMatrix(i, .ColIndex("NoteSerial1"))) & ","
                StrSQL = StrSQL & SqlNum(val(.TextMatrix(i, .ColIndex("Note_Value")))) & ","
                StrSQL = StrSQL & SqlNum(val(.TextMatrix(i, .ColIndex("PayedValue")))) & ","
                StrSQL = StrSQL & "'" & SqlQ(.TextMatrix(i, .ColIndex("too"))) & "',"
                StrSQL = StrSQL & SqlDateOrNull(.TextMatrix(i, .ColIndex("DueDate"))) & ","
                StrSQL = StrSQL & SqlDateOrNull(.TextMatrix(i, .ColIndex("NoteDate"))) & ","
                StrSQL = StrSQL & SqlNum(Diff) & ","
                StrSQL = StrSQL & SqlNum(val(.TextMatrix(i, .ColIndex("NetValue")))) & ","
                StrSQL = StrSQL & SqlNum(Remaining)
                StrSQL = StrSQL & ")"

                Cn.Execute StrSQL, , adExecuteNoRecords

                '================================================
                '  ÕœÌÀ Õ«·… «·›« Ê—… („œ›Ê⁄… »«·ﬂ«„· √„ ·«)
                '================================================
                If val(.TextMatrix(i, .ColIndex("NetValue"))) <= 0 Then
                    StrSQL = "UPDATE Transactions SET TotalPayed=1 WHERE Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
                Else
                    StrSQL = "UPDATE Transactions SET TotalPayed=0 WHERE Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
                End If
                Cn.Execute StrSQL, , adExecuteNoRecords

            End If

ContinueLoop:
        Next i

    End With

    '========================================================
    ' 4) ”Ã· ›Ì TblBillBuyPayment (“Ì „« ﬂ«‰ ‘€«· ⁄‰œﬂ)
    '========================================================
    With VSFlexGrid1
        For i = .FixedRows To .rows - 1

            If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then

                Diff = val(.TextMatrix(i, .ColIndex("TransPayedValue")))

                If Diff > 0 Then
                    StrSQL = "INSERT INTO TblBillBuyPayment (NoteID, Transaction_ID, Note_Value, PayedValue) VALUES (" & _
                             val(Me.XPTxtID.text) & "," & _
                             val(.TextMatrix(i, .ColIndex("NoteID"))) & "," & _
                             SqlNum(val(.TextMatrix(i, .ColIndex("Note_Value")))) & "," & _
                             SqlNum(Diff) & ")"
                    Cn.Execute StrSQL, , adExecuteNoRecords
                End If

            End If
        Next i
    End With

 '   Cn.CommitTrans
    saveBillBuy = True
    Exit Function

EH:
    On Error Resume Next
 '   Cn.RollbackTrans
    MsgBox "Error in saveBillBuy: " & Err.Number & " - " & Err.Description, vbExclamation
    saveBillBuy = False
End Function


'
Function saveBillBuyOLD()
    Dim StrSQL As String
   ' Dim StrSQL  As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Diff = 0
Dim RsDetails As ADODB.Recordset
      If Me.TxtModFlg.text = "E" Then
    StrSQL = "Delete From TblNotesBillBuyPayment Where NoteID1=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblNotesBillBuyPayment Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
    TxtValueTemp.text = val(XPTxtVal.text)
    For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID1").value = val(XPTxtID.text)
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
            RsDetails("DueDate").value = IIf(.TextMatrix(i, .ColIndex("DueDate")) = "", Null, .TextMatrix(i, .ColIndex("DueDate")))
            RsDetails("NoteDate").value = (.TextMatrix(i, .ColIndex("NoteDate")))
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

    StrSQL = "SELECT     * from dbo.TblBillBuyPayment Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
    For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.text)
            RsDetails("Transaction_ID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails.update
        End If
    Next i
End With

End Function
Function saveBillProject()
    Dim StrSQL As String
   ' Dim StrSQL  As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Diff = 0
Dim RsDetails As ADODB.Recordset
      If Me.TxtModFlg.text = "E" Then
      
    StrSQL = "Delete From TblNotesBillProjectPayment Where NoteID1=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillProjectPayment  Where NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblNotesBillProjectPayment Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid2
    TxtValueTemp.text = val(XPTxtVal.text)
    For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID1").value = val(XPTxtID.text)
            RsDetails("project_no").value = val(.TextMatrix(i, .ColIndex("project_no")))
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
            RsDetails("NoteDate").value = (.TextMatrix(i, .ColIndex("NoteDate")))
            RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
          .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
             RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            RsDetails.update
                
            If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
            StrSQL = "Update project_billl Set  TotalPayed=1 Where id=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
             Else
                 StrSQL = "Update project_billl Set  TotalPayed=0 Where id=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
      End If
    Next i
End With
    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblBillProjectPayment  Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid2
    For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.text)
            RsDetails("Transaction_ID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("project_no").value = val(.TextMatrix(i, .ColIndex("project_no")))
            RsDetails.update
        End If
    Next i
End With

End Function
Function saveBillVendor()
    Dim StrSQL As String
   ' Dim StrSQL  As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Diff = 0
Dim RsDetails As ADODB.Recordset
      If Me.TxtModFlg.text = "E" Then
    StrSQL = "Delete From TblNotesBillVindorPayment Where NoteID1=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillVindorPayment Where NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblNotesBillVindorPayment Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With GRID1
    TxtValueTemp.text = val(XPTxtVal.text)
    For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID1").value = val(XPTxtID.text)
            
            RsDetails("InstalValue").value = val(.TextMatrix(i, .ColIndex("InstalValue")))
            RsDetails("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
            RsDetails("branch_no").value = val(.TextMatrix(i, .ColIndex("branch_no")))
            RsDetails("NoteSerial1").value = val(.TextMatrix(i, .ColIndex("NoteSerial1")))
            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
            Note_Value1 = val(.TextMatrix(i, .ColIndex("RemainingValue")))
          '  RsDetails("QestID").value = val(.TextMatrix(i, .ColIndex("QestID")))
            RsDetails("StrQest").value = (.TextMatrix(i, .ColIndex("StrQest")))
            If Aut_manual = False Then
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
            End If
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
            
            RsDetails("too").value = (.TextMatrix(i, .ColIndex("too")))
            RsDetails("NoteDate").value = (.TextMatrix(i, .ColIndex("NoteDate")))
            RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
          .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
             RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            RsDetails.update
            If (.TextMatrix(i, .ColIndex("StrQest"))) <> "" Then
                     StrSQL = "Update TblQestFexed Set  FlgPaye=1  Where (QestID in(" & (.TextMatrix(i, .ColIndex("StrQest"))) & "))  and Ind=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                End If
                   StrSQL = "Update notes_all Set  FlgPaye=1  Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
                
            If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
            StrSQL = "Update notes_all Set  TotalPayed=1 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
             Else
                 StrSQL = "Update notes_all Set  TotalPayed=0 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
        End If
    Next i
End With
    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblBillVindorPayment Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With GRID1
    For i = .FixedRows To .rows - 1
        If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.text)
            RsDetails("NoteSerial1").value = val(.TextMatrix(i, .ColIndex("NoteSerial1")))
            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails.update
        End If
    Next i
End With

End Function
Function saveChequeBoxContents1(NoteID As Double)
    If SystemOptions.IsCheque = True And CboPayMentType.ListIndex = 1 Then
    Else
        If SystemOptions.banks_Accounts3 = False Then Exit Function
    End If
    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords
 
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
        rs("ChequeValue").value = val(XPTxtVal.text) + val(TxtPrePayd(17).text)
 '   rs("ChequeValue").value = val(XPTxtVal.Text) + val(TxtVAt2.Text)
        rs("Remarks").value = Me.DcboDebitSide.text
        rs("Payed").value = 0
       
        rs("DepitAccount").value = (DcboDebitSide.BoundText)
        rs.update
    End If

    rs.Close
End Function

Function save_cost_center()

    'on error resume next
    If Not IsNumeric(Text1.text) Then Exit Function
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
        rs("NoteSerial").value = TxtNoteSerial.text
        rs("Remark").value = "”‰œ  ’—›  „œ›Ê⁄«  —ﬁ„ " & TxtNoteSerial1 & "    " & Me.txt_general_des
 
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

    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND kedno =" & val(Text1.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
 
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If
        
    'rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  StrSQL = "SELECT   *  from dbo.marakes_taklefa_temp Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    'ÿ—› „œÌ‰
    rs.AddNew
    rs("general_des").value = 1
    rs("cost_center_id").value = cost_center_id
    rs("cost_center").value = cost_center
    rs("value").value = XPTxtVal.text
    rs("depit_or_credit").value = "„œÌ‰"
    rs("opr_id").value = Me.Text1.text
    rs("kedno").value = Me.Text1.text
        
    rs("opr_type").value = opr_type
    rs("account_name").value = DcboDebitSide.text
    rs("account_no").value = DcboDebitSide.BoundText
    rs("line_no").value = Line1
    rs("record_date").value = record_date
                    rs("description").value = txt_general_des.text
                    
    rs.update
    'ÿ—› œ«∆‰
    '    rs.AddNew
    '    rs("cost_center_id").value = cost_center_id
    '    rs("cost_center").value = cost_center
    '    rs("value").value = XPTxtVal.text
    '    rs("depit_or_credit").value = "œ«∆‰"
    '    rs("opr_id").value = Me.Text1.text
    '    rs("kedno").value = Me.Text1.text
    '
    '    rs("opr_type").value = opr_type
    '    rs("account_name").value = DcboCreditSide.text
    '    rs("account_no").value = DcboCreditSide.BoundText
    '    rs("line_no").value = Line2
    '    rs("record_date").value = record_date
    '    rs.update
 
    rs.Close
End Function

Function FIFO_FUNCTION(CusID As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    sql = "SELECT CompanyDebitValues.* FROM dbo.CompanyDebitValues() CompanyDebitValues  where   (cusid=" & CusID & " and requiredvalue>0)"

    'Sql = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues  where   (cusid=" & CusID & " and requiredvalue>0)"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Dim total_value As Double
    Dim current_value As Double
    total_value = val(txtAdv_payment_value.text)
  
    For i = 1 To Rs3.RecordCount

        If total_value > Rs3("requiredvalue") Then
            current_value = Rs3("requiredvalue")
            total_value = total_value - current_value
        
        ElseIf total_value <= Rs3("requiredvalue") Then
            current_value = total_value
            total_value = 0
        ElseIf total_value = 0 Then
            Exit Function
        End If
  
        Add_new_notes Me.XPDtbTrans, 2001, current_value, Rs3("transactionsid").value, CusID, val(DcboBox.BoundText), 1, val(DCboUserName.BoundText)
  
        Rs3.MoveNext
    Next i

    ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
    txtAdv_payment_value.text = total_value
    change_adv_payment_value XPTxtID.text, total_value
    ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
    Rs3.Close

End Function

Function Add_new_notes(NoteDate As Date, NoteType As Integer, Note_Value As Double, Transaction_ID As Double, CusID As Double, BoxID As Integer, displayed As Integer, UserID As Integer)
    Dim RsDev As New ADODB.Recordset
    Exit Function
 '   RsDev.Open "notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Dim StrSQL As String

    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
    RsDev.AddNew
      
    RsDev("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    RsDev("NoteSerial").value = TxtNoteSerial.text ' CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=2000"))
              
    RsDev("NoteDate").value = NoteDate
    RsDev("NoteType").value = NoteType
           
    RsDev("Note_Value").value = Note_Value
    RsDev("Transaction_ID").value = Transaction_ID
    RsDev("CusID").value = CusID
    RsDev("BoxID").value = BoxID
    RsDev("UserID").value = UserID
    RsDev("displayed").value = 0
           
    RsDev.update

End Function

Function change_adv_payment_value(note_id As Double, value As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer

    sql = "SELECT * from notes   where  NoteID=" & note_id
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Rs3("Adv_payment_value").value = value
    Rs3.update
  
End Function

Function Distribute_to_bills(SQL1 As String, CusID As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer

    sql = "SELECT CompanyDebitValues.* FROM dbo.CompanyDebitValues() CompanyDebitValues  where    requiredvalue>0 and " & SQL1

    'Sql = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues  where  requiredvalue>0 and " & Sql1
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Dim total_value As Double
    Dim current_value As Double
    total_value = val(txtAdv_payment_value.text)
  
    For i = 1 To Rs3.RecordCount

        If total_value > Rs3("requiredvalue") Then
            current_value = Rs3("requiredvalue")
            total_value = total_value - current_value
        
        ElseIf total_value <= Rs3("requiredvalue") Then
            current_value = total_value
            total_value = 0
        ElseIf total_value = 0 Then
            Exit Function
        End If
  
        Add_new_notes Me.XPDtbTrans, 2001, current_value, Rs3("transactionsid").value, CusID, DcboBox.BoundText, 1, val(DCboUserName.BoundText)
        Rs3.MoveNext
    Next i

    txtAdv_payment_value.text = total_value
    change_adv_payment_value XPTxtID.text, total_value

    ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
  
    ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
    Rs3.Close
 
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
                Msg = Msg & CHR(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  ”ÃÌ· „œ›Ê⁄«  ·Â«"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ : " & Trim(Me.TxtTransSerial.text)
                Msg = Msg & CHR(13) & "—ﬁ„ «·›« Ê—… ›Ï «·»—‰«„Ã : " & Me.TxtTransID.text
                Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· : " & Me.DBCboClientName.text
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
                        Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘… «·„œ›Ê⁄« "
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
            StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & "Notes.Transaction_ID " & " Where ((Notes.NoteType = 5 OR Notes.NoteType = 10) And Transactions.Transaction_ID = " & LngTransID & ")"

            If Me.TxtModFlg.text = "E" Then
                StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.text & ""
            End If

            StrSQL = StrSQL + " GROUP BY Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType "
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If DblCreditNoteValue = RsTemp("SumNote_Value").value Then
                    Msg = "⁄›Ê« ...!!!!!" & CHR(13)
                    Msg = Msg & "·ﬁœ  „  ”ÃÌ· „œ›Ê⁄«  √Ê (⁄„· Œ’Ê„«  „ﬂ ”»…) ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
                    Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰   ”ÃÌ· «Ì… „œ›Ê⁄«  ≈÷«›Ì… ⁄·ÌÂ«."
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Function
                ElseIf RsTemp("SumNote_Value").value + val(Me.XPTxtVal.text) > DblCreditNoteValue Then
                    Msg = "⁄›Ê« ..."
                    Msg = Msg & CHR(13) & "·ﬁœ  „  ”ÃÌ· „œ›Ê⁄«  √Ê (⁄„· Œ’Ê„«  „ﬂ ”»…) „”»ﬁ« ·Â–Â «·›« Ê—…"
                    Msg = Msg & CHR(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
                    Msg = Msg & CHR(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
                    Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                    Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                    Msg = Msg & CHR(13) & "ﬁÌ„… «·„œ›Ê⁄«  «Ê «·Œ’Ê„«  «·„ﬂ ”»… «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").value
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
Sub DeleteBillProject()
Dim i As Integer
Dim StrSQL As String
With VSFlexGrid2
 For i = .FixedRows To .rows - 1
 If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
      StrSQL = "Update project_billl Set  TotalPayed=0 Where Id=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
     End If
     Next i
 End With
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


Sub DeleteBill()
Dim i As Integer
Dim StrSQL As String
With GRID1
 For i = .FixedRows To .rows - 1
 If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
      StrSQL = "Update notes_all Set  TotalPayed=0 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
     End If
     Next i
 End With
End Sub
Private Sub Del_Trans()
    Dim Msg As String
     On Error GoTo ErrTrap

    If SystemOptions.banks_Accounts3 = True Then
        If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
            Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–… «·⁄„·Ì…"
            Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    
    End If
         If CheAssetPayd(val(Me.XPTxtID)) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–Â «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ≈÷«›… ··«’Ê·   "
                    Else
                    Msg = " Can Not Delete this Process"
                    Msg = Msg & CHR(13) & " There is the Process of adding Assest "
                    
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
                
    If XPTxtID.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    Else
        Msg = Msg + " Confrim Delete ?"
     
End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then

Dim My_SQL As String
If PayDes.text <> "" Then

My_SQL = " update TblPripaidExpensesDet Set PaymentPayed = 0  Where   (id in (" & PayDes.text & "))"
                  
Cn.Execute My_SQL
End If

                CuurentLogdata ("D")
DeletePayedSalary Me.CboYear1.text & CmbMonth1.ListIndex + 1, empDes
Me.DeletePayedPayment PayDes.text
Me.DeletePayedPayment2 TxtNoSupplerDes.text
DeletePayedPaymeQest
 If val(DCboCashType.ListIndex) = 8 Then
  Cn.Execute "update TblVocationEntitlements set PayedPayment =Null where ID=" & val(TxtDue.text) & ""
  End If
   If val(DCboCashType.ListIndex) = 12 Then
  Cn.Execute "update TblVATAvowal set Paid =Null where ID=" & val(TxtEndService.text) & ""
  End If

                rs.delete
                Dim StrSQL As String
       
           Cn.Execute "update TblEmpAdvanceRequest set AccAproved=Null  where AdvanceID=" & val(TxtAdvance.text) & ""
            '    StrSQL = "Delete From notes  Where (NoteType=2001 OR NoteType=5 ) AND NoteSerial=" & val(TxtNoteSerial.Text)
            '    Cn.Execute StrSQL, , adExecuteNoRecords
        Cn.Execute "Delete from TblSalaryNotesPayment where TransID=" & val(XPTxtID.text) & ""
                StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
       
                StrSQL = "Delete From ReciveDetails Where NoteSerial1='" & val(TxtNoteSerial1.text) & "'"
                Cn.Execute StrSQL, , adExecuteNoRecords
    
                StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
    
                StrSQL = "Delete From TblEmpAdvance Where AdvanceID=" & val(Me.XPTxtID1.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
         If val(DCboCashType.ListIndex) = 10 Then
          StrSQL = "Update End_of_service set PaymPaid=null where id=" & val(Me.TxtEndService.text) & " "
            Cn.Execute StrSQL
            End If
                StrSQL = "Delete From TblEmpAdvanceDetails Where AdvanceID=" & val(Me.XPTxtID1.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                DeleteBill
                DeleteBillBuy
                DeleteBillProject
                             StrSQL = "Delete From TblNotesBillProjectPayment Where NoteID1=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillProjectPayment  Where NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
             StrSQL = "Delete From TblNotesBillVindorPayment Where NoteID1=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillVindorPayment Where NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
                 StrSQL = "Delete From TblNotesBillBuyPayment Where NoteID1=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment Where TypTrans IS NULL and NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If

                WriteInfo
            End If
        End If

    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This Process is not Available does not Have any Record"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry an Error occurred during the Deletion " & CHR(13)
 End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
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

    MySQL = "Select * From Expanses_Order  where ChqueNum='0'"

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

    '
    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If RsData.BOF Or RsData.EOF Then
    'GetMsgs 138, vbExclamation
    '    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    RsData.Close
    '    Set RsData = Nothing
    '    Screen.MousePointer = vbDefault
    '    Exit Function
    'End If
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
    xReport.ParameterFields(11).AddCurrentValue CStr(XPMTxtRemarks.text)
'    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtValView.text)
If right(XPTxtValView, 2) = "00" Then
    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtVal.text)
    Else
    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtValView.text)
    End If
    
    xReport.ParameterFields(13).AddCurrentValue CStr(Me.txtperson.text)
    xReport.ParameterFields(14).AddCurrentValue CStr(lbl(18).Caption)
    xReport.ParameterFields(15).AddCurrentValue Format$(DtpChequeDueDate.value, "dd/mm/yyyy")
 
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function
 
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·”‰œ " & TxtNoteSerial1.text & CHR(13) & "   «· «—ÌŒ " & XPDtbTrans & CHR(13) & "   ‰Ê⁄ «·„œ›Ê⁄«  " & DCboCashType & CHR(13) & "   «·›—⁄  " & dcBranch & CHR(13) & "   «·«”„  " & DBCboClientName & CHR(13) & "   ﬁÌ„Â «·„œ›Ê⁄«   " & XPTxtVal & CHR(13) & "   ÿ—Ìﬁ… «·œ›⁄ " & CboPayMentType & CHR(13) & "   «·Œ“Ì‰…  " & DcboBox & CHR(13) & "   «·»‰ﬂ  " & DcboBankName & CHR(13) & "   —ﬁ„ «·‘Ìﬂ  " & TxtChequeNumber & CHR(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & CHR(13) & "     »‰«¡ ⁄·Ï   " & XPMTxtRemarks & CHR(13) & "   «·‘—Õ «·⁄«„    " & txt_general_des & CHR(13) & "     —ﬁ„ «·ÿ·»Ì…  " & TXT_order_no & CHR(13) & "  „—ﬂ“ «· ﬂ·›… «·⁄«„  " & DcCostCenter & CHR(13) & "   —”Ê„ «·ÕÊ«·… " & txtTransferExpenses & CHR(13) & "   —ﬁ„ «·ﬁÌœ   " & TxtNoteSerial & CHR(13) & "ÿ—› „œÌ‰  " & DcboDebitSide & CHR(13) & " ÿ—› œ«∆‰ " & DcboCreditSide & CHR(13) & "«”„ «·„” Œœ„ " & DCboUserName
                        
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Vchr. NO.  " & TxtNoteSerial1.text & CHR(13) & "   Date " & XPDtbTrans & CHR(13) & "  Payment Type " & DCboCashType & CHR(13) & "   Branch  " & dcBranch & CHR(13) & "   Name  " & DBCboClientName & CHR(13) & "  Value" & XPTxtVal & CHR(13) & "   Cash/   Cheque " & CboPayMentType & CHR(13) & "   Box  " & DcboBox & CHR(13) & "   Bank  " & DcboBankName & CHR(13) & "   Cheque No" & TxtChequeNumber & CHR(13) & "  Due Date  " & DtpChequeDueDate & CHR(13) & "  Based On " & XPMTxtRemarks & CHR(13) & "  General Des  " & txt_general_des & CHR(13) & " Order No " & TXT_order_no & CHR(13) & " Cost Center " & DcCostCenter & CHR(13) & "  Transfer Cost " & txtTransferExpenses & CHR(13) & " Ge NO.  " & TxtNoteSerial & CHR(13) & "Debit " & DcboDebitSide & CHR(13) & "Credit " & DcboCreditSide & CHR(13) & " UserName " & DCboUserName
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 5, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 5, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtNoteSerial, TxtNoteSerial1
    End If
    
End Function
Function ChekExpens(Optional PayDes As String) As Boolean
Dim sql As String
If PayDes <> "" Then
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "SELECT     dbo.TblPripaidExpensesDet.ID, dbo.TblPripaidExpChiled.PaidExIDDet, dbo.TblPripaidExpChiled.Etfa"
sql = sql & " FROM         dbo.TblPripaidExpensesDet LEFT OUTER JOIN"
sql = sql & "                      dbo.TblPripaidExpChiled ON dbo.TblPripaidExpensesDet.ID = dbo.TblPripaidExpChiled.PaidExIDDet"
sql = sql & "  WHERE     (dbo.TblPripaidExpChiled.Etfa = 1) AND (dbo.TblPripaidExpensesDet.ID IN (" & PayDes & "))"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekExpens = True
Else
ChekExpens = False
End If
Else
End If
End Function
Function ChekExpensTotal(Optional PayDes As String) As Boolean
Dim sql As String
If PayDes <> "" Then
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
'sql = "SELECT     ID, PaymentPayed"
'sql = sql & " From dbo.TblPripaidExpensesDet"
'sql = sql & " WHERE     (PaymentPayed = 1) AND (ID IN (" & PayDes & "))"

sql = "select * from TblPripaidExpChiled where  etfa=1 and paidexiddet in (" & PayDes & ")  "

Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekExpensTotal = True
Else
ChekExpensTotal = False
End If
Else
ChekExpensTotal = False
End If
End Function
Function CheAdvanced(Optional advanceID As Double = 0) As Boolean
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
CheAdvanced = False
sql = "select AdvanceID from TblEmpAdvancePayedDet where AdvanceID=" & advanceID & ""
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CheAdvanced = True
Else
CheAdvanced = False
End If
End Function

Function CheAssetPayd(Optional NoteID As Double = 0) As Boolean
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
CheAssetPayd = False
sql = "select NoteID from Notes where NoteID=" & NoteID & " and (AssestPayd =1) "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CheAssetPayd = True
Else
CheAssetPayd = False
End If
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
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(14), "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

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
       
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPDtbTrans_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
    If CheckAnyVAT(XPDtbTrans.value) = False Then
IncludVAT.value = vbUnchecked
IncludVAT.Enabled = False
Else
IncludVAT.Enabled = True
End If
End Sub

Private Sub Txt_DateHigri_LostFocus()
    XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)
 
End Sub

Private Sub XPTxtID_Change()
If called = True Then Exit Sub
    If Me.TxtModFlg <> "N" And Me.TxtModFlg <> "E" Then
        DCboCashType_Change
    End If

End Sub

Private Sub XPTxtID1_Change()

    If Me.TxtModFlg = "R" Or Me.TxtModFlg = "" Then
        If val(XPTxtID1.text) <> 0 Then
            Fra(2).Visible = True
             lbl(47).Visible = True
        TxtAdvance.Visible = True
            
        Else
            Fra(2).Visible = False
             lbl(47).Visible = False
        TxtAdvance.Visible = False
        End If
    End If
    If (Me.TxtModFlg = "E" Or Me.TxtModFlg = "N") Or val(XPTxtID1.text) = 0 Then Exit Sub
        
        RetriveAdvanced val(XPTxtID1.text)
   
End Sub
Public Sub RetriveBillBuyData(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String


   ' On Error GoTo ErrTrap
    Set RsDetails = New ADODB.Recordset
  StrSQL = "   SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblNotesBillBuyPayment.*"
  StrSQL = StrSQL & "  FROM         dbo.TblNotesBillBuyPayment LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblNotesBillBuyPayment.branch_no = dbo.TblBranchesData.branch_id"
  StrSQL = StrSQL & "  Where (dbo.TblNotesBillBuyPayment.NoteID1 = " & val(XPTxtID.text) & ")"
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid1
    .Clear flexClearScrollable, flexClearEverything
    .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
      '  Fra(2).Visible = True
      '               lbl(47).Visible = True
      '  TxtAdvance.Visible = True
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
            .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(RsDetails("DueDate").value), "", RsDetails("DueDate").value)
            .TextMatrix(i, .ColIndex("RemainingValue")) = IIf(IsNull(RsDetails("RemainingValue").value), 0, RsDetails("RemainingValue").value)
           ' .TextMatrix(i, .ColIndex("PartValue")) = Round(RsDetails("PartValue").value, 2)
            .TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(CDate(RsDetails("NoteDate").value))
            .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
            RsDetails.MoveNext
        Next i
        

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
End Sub
Public Sub RetriveBillProjectData(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String


   ' On Error GoTo ErrTrap
    Set RsDetails = New ADODB.Recordset
  StrSQL = " SELECT     dbo.TblNotesBillProjectPayment.ID, dbo.TblNotesBillProjectPayment.NoteID, dbo.TblNotesBillProjectPayment.NoteSerial1, "
  StrSQL = StrSQL & "                    dbo.TblNotesBillProjectPayment.PayedValue, dbo.TblNotesBillProjectPayment.too, dbo.TblNotesBillProjectPayment.Note_Value,"
  StrSQL = StrSQL & "                    dbo.TblNotesBillProjectPayment.NoteDate, dbo.TblNotesBillProjectPayment.RemainingValue, dbo.TblNotesBillProjectPayment.TransPayedValue,"
  StrSQL = StrSQL & "                    dbo.TblNotesBillProjectPayment.NetValue, dbo.TblNotesBillProjectPayment.branch_no, dbo.TblNotesBillProjectPayment.NoteID1,"
  StrSQL = StrSQL & "                    dbo.TblNotesBillProjectPayment.project_no, dbo.projects.Fullcode, dbo.projects.Project_name, dbo.projects.Project_nameE, dbo.TblBranchesData.branch_name,"
  StrSQL = StrSQL & "                    dbo.TblBranchesData.branch_namee  "
  StrSQL = StrSQL & " FROM         dbo.TblNotesBillProjectPayment INNER JOIN"
  StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.TblNotesBillProjectPayment.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.projects ON dbo.TblNotesBillProjectPayment.project_no = dbo.projects.id"
  StrSQL = StrSQL & "  Where (dbo.TblNotesBillProjectPayment.NoteID1 = " & val(XPTxtID.text) & ")"
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid2
    .Clear flexClearScrollable, flexClearEverything
    .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To RsDetails.RecordCount
        .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("project_no")) = IIf(IsNull(RsDetails("project_no").value), 0, RsDetails("project_no").value)
            .TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(RsDetails("branch_no").value), 0, RsDetails("branch_no").value)
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(RsDetails("Project_name").value), "", RsDetails("Project_name").value)
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
            Else
            .TextMatrix(i, .ColIndex("Project_name")) = IIf(IsNull(RsDetails("Project_nameE").value), "", RsDetails("Project_nameE").value)
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
            '.TextMatrix(i, .ColIndex("RemainingValueE")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) * IIf(val(.TextMatrix(i, .ColIndex("Currency_rate"))) <> 0, val(.TextMatrix(i, .ColIndex("Currency_rate"))), 1)
           ' .TextMatrix(i, .ColIndex("PartValue")) = Round(RsDetails("PartValue").value, 2)
            .TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(CDate(RsDetails("NoteDate").value))
            .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
            RsDetails.MoveNext
        Next i
        

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
End Sub
Public Sub RetriveBillVendorData(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String


   ' On Error GoTo ErrTrap
    Set RsDetails = New ADODB.Recordset
  StrSQL = "   SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblNotesBillVindorPayment.*"
  StrSQL = StrSQL & "  FROM         dbo.TblNotesBillVindorPayment LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblNotesBillVindorPayment.branch_no = dbo.TblBranchesData.branch_id"
  StrSQL = StrSQL & "  Where (dbo.TblNotesBillVindorPayment.NoteID1 = " & val(XPTxtID.text) & ")"
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With GRID1
    .Clear flexClearScrollable, flexClearEverything
    .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
      '  Fra(2).Visible = True
      '               lbl(47).Visible = True
      '  TxtAdvance.Visible = True
        RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To RsDetails.RecordCount
        .TextMatrix(i, .ColIndex("Ser")) = i
        .TextMatrix(i, .ColIndex("StrQest")) = IIf(IsNull(RsDetails("StrQest").value), 0, RsDetails("StrQest").value)
       .TextMatrix(i, .ColIndex("InstalValue")) = IIf(IsNull(RsDetails("InstalValue").value), 0, RsDetails("InstalValue").value)

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
            
           ' .TextMatrix(i, .ColIndex("PartValue")) = Round(RsDetails("PartValue").value, 2)
            .TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(CDate(RsDetails("NoteDate").value))
            .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
            RsDetails.MoveNext
        Next i
        

    End If
End With
    RsDetails.Close
    Set RsDetails = Nothing
    Set rs = Nothing
Exit Sub
ErrTrap:
End Sub
Public Sub RetriveAdvanced(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String
    If Lngid = 0 Then Exit Sub
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEmpAdvance  Where (TblEmpAdvance.AdvanceType =0) Order By AdvanceID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount < 1 Then
        Fra(2).Visible = False
         lbl(47).Visible = False
        TxtAdvance.Visible = False
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Fra(2).Visible = False
                     lbl(47).Visible = False
        TxtAdvance.Visible = False
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "AdvanceID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID1.text = IIf(IsNull(rs("AdvanceID").value), "", val(rs("AdvanceID").value))
    'XPDtbTrans.value = IIf(IsNull(rs("AdvanceDate").value), Date, rs("AdvanceDate").value)

    'Me.DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)

    XPTxtVal.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
    'Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)

    Me.TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), "", rs("PaymentCounts").value)
    Me.CmbMonth.ListIndex = rs("FirstMonthPayment").value - 1
    Me.CboYear.text = rs("FirstYearPayment").value
    Me.ChkSaleryDis.value = IIf(rs("AutoDiscount").value = True, vbChecked, vbUnchecked)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)

    Set RsDetails = New ADODB.Recordset
    StrSQL = "Select * From  TblEmpAdvanceDetails Where AdvanceID=" & val(XPTxtID1.text)
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = FG.FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        Fra(2).Visible = True
                     lbl(47).Visible = True
        TxtAdvance.Visible = True
        RsDetails.MoveFirst
        FG.rows = FG.FixedRows + RsDetails.RecordCount

        For i = Me.FG.FixedRows To FG.rows - 1
            FG.TextMatrix(i, FG.ColIndex("PartNO")) = RsDetails("PartNO").value
            FG.TextMatrix(i, FG.ColIndex("PartValue")) = Round(RsDetails("PartValue").value, 2)
            FG.TextMatrix(i, FG.ColIndex("PartDate")) = DisplayDate(CDate(RsDetails("PartDate").value))
            RsDetails.MoveNext
        Next i
        

    End If

    RsDetails.Close
    Set RsDetails = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
End Sub
Sub CalcuteValue()
    Dim NotValue As Double
    
If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
If (TxtOrder.text <> "") Then
            NotValue = GetVal((Me.TxtOrder.text), val(Me.XPTxtID), 5)
            If Price < NotValue + val(Me.XPTxtVal) Then
            XPTxtVal = ""
            MsgBox "€Ì— „”„ÊÕ »Â–… «·ﬁÌ„Â ·«‰Â«   ŒÿÌ «·’—›  «·ﬁÌ„Â «·„ »ﬁÌÂ ÂÌ " & Price - NotValue, vbCritical
            Exit Sub
            End If
End If

End If

 
     'Me.lbl(18).Caption = WriteNo(Me.XPTxtVal.text, 0, True)
    XPTxtValView.text = Format(val(XPTxtVal.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
   ' XPTxtVal.Text = Format(XPTxtVal.Text, "###.00")
   If XPTxtVal <> "" Then
   
Text4.text = Format(XPTxtVal.text, "###.00")
Else
Text4 = 0
End If
If val(TxtPrePayd(17).text) > 0 And (val(DCboCashType.ListIndex) = 1 Or val(DCboCashType.ListIndex) = 2 Or val(DCboCashType.ListIndex) = 0 Or val(DCboCashType.ListIndex) = 5 Or val(DCboCashType.ListIndex) = 7) And Option3.value = True Then
Text4 = val(Text4.text) + val(TxtPrePayd(17).text)
End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(18).Caption = WriteNo(Text4.text, 0, True, ".", , 0)

    Else
 
         Me.lbl(18).Caption = WriteNo(Text4.text, 0, True, ".", , 1)

    End If

    If TxtModFlg.text = "N" Then
        txtAdv_payment_value.text = XPTxtVal.text
    End If
End Sub
Private Sub XPTxtVal_Change()
'txtTotalWithVat = 0
CalcuteValue


  
CalCuteCurrency
End Sub

Private Sub XPTxtVal_GotFocus()
'XPTxtVal.Text = Format(XPTxtVal.Text, "###.00")
End Sub

Private Sub XPTxtVal_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
           
           LostAllFocus
        End If
End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
  '  CalCulteVAT
End Sub

Private Sub WriteInfo()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StartWeekDate As Date
    Dim EndWeekDate As Date
    Dim StrTemp As String
    Dim i As Integer

    StartWeekDate = GetWeekStartEND(Date, 0)
    EndWeekDate = DateAdd("d", 7, StartWeekDate)

    If SystemOptions.UserInterface = ArabicInterface Then
        StrTemp = "«·≈”»Ê⁄ «·Õ«·Ï „‰ " & DisplayDate(StartWeekDate)
        StrTemp = StrTemp & " ≈·Ï " & DisplayDate(EndWeekDate)
    Else
        StrTemp = " Current Week From" & DisplayDate(StartWeekDate)
        StrTemp = StrTemp & " To " & DisplayDate(EndWeekDate)

    End If

    Me.lbl(22).Caption = StrTemp

    For i = LblLinkInfo.LBound To LblLinkInfo.UBound
        LblLinkInfo(i).Caption = "0"
    Next i

    '------------------------------------------------------------------------------
    '„œ›Ê⁄«  «·ÌÊ„
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 5) "
    StrSQL = StrSQL + " AND NoteDate=" & SQLDate(Date, True)
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(0).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(1).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(6).Caption = val(Me.LblLinkInfo(0).Caption) + val(Me.LblLinkInfo(1).Caption)
    Else
        Me.LblLinkInfo(0).Caption = 0
        Me.LblLinkInfo(1).Caption = 0
        Me.LblLinkInfo(6).Caption = 0
    End If

    '------------------------------------------------------------------------------
    '„œ›Ê⁄«  «·√”»Ê⁄ «·Õ«·Ï
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 5) "
    StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(StartWeekDate, True)
    StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(EndWeekDate, True)
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(2).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(3).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(7).Caption = val(Me.LblLinkInfo(2).Caption) + val(Me.LblLinkInfo(3).Caption)
    Else
        Me.LblLinkInfo(0).Caption = 0
        Me.LblLinkInfo(1).Caption = 0
        Me.LblLinkInfo(7).Caption = 0
    End If

    '------------------------------------------------------------------------------
    '„œ›Ê⁄«  «·‘Â— «·Õ«·Ï
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 5) "
    StrSQL = StrSQL + " AND Month(NoteDate)=" & Month(Date) & ""
    StrSQL = StrSQL + " AND Year(NoteDate)=" & year(Date) & ""
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(4).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(5).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(8).Caption = val(Me.LblLinkInfo(4).Caption) + val(Me.LblLinkInfo(5).Caption)
    Else
        Me.LblLinkInfo(4).Caption = 0
        Me.LblLinkInfo(5).Caption = 0
        Me.LblLinkInfo(8).Caption = 0
    End If

End Sub

Private Sub XPTxtVal_KeyUp(KeyCode As Integer, Shift As Integer)
CalCuteCurrency
End Sub

Private Sub XPTxtVal_LostFocus()
LostAllFocus
If val(DCboCashType.ListIndex) = 1 Then
If val(Label16.Caption) > 0 Then
If val(XPTxtVal.text) > val(Label16.Caption) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ﬁÌ„… «·„œ›Ê⁄«  «ﬂ»— „‰ „Ã„Ê⁄ ﬁÌ„ «·›Ê« Ì— "
Else
MsgBox "Can Not Value of Payment Larger than Total of Bills "
End If
XPTxtVal.text = 0
XPTxtVal.SetFocus
Exit Sub
End If
End If
End If
End Sub

Function LostAllFocus()
 
CalCuteCurrencyE
 'XPTxtVal.Text = Format(XPTxtVal.Text, "#,##0.00")
'XPTxtValE.Text = Format(XPTxtValE.Text, "#,##0.00")
End Function

Private Sub XPTxtVal_Validate(Cancel As Boolean)
ClaCul
End Sub

Private Sub XPTxtValE_GotFocus()
'XPTxtValE.Text = Format(XPTxtValE.Text, "###.00")
End Sub

Private Sub XPTxtValE_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
           LostAllFocus
        End If
End Sub

Private Sub XPTxtValE_LostFocus()
LostAllFocus
End Sub



Private Sub XPTxtValE_Validate(Cancel As Boolean)
CalCuteCurrencyE
End Sub
