VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmPayments2 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "”‰œ ’—› - «·„œ›Ê⁄«   "
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14325
   HelpContextID   =   390
   Icon            =   "FrmPayments2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   14325
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
   Begin VB.OptionButton ComResid 
      Alignment       =   1  'Right Justify
      Caption         =   " Ã«—Ì"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   230
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton ComResid 
      Alignment       =   1  'Right Justify
      Caption         =   "”ﬂ‰Ì"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   229
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H0080FFFF&
      Caption         =   "»Ì«‰«  „ﬁ»Ê÷«  «·⁄ﬁ«—« "
      Height          =   6855
      Index           =   1
      Left            =   -390
      RightToLeft     =   -1  'True
      TabIndex        =   210
      Top             =   600
      Visible         =   0   'False
      Width           =   14415
      Begin VB.Frame Frame11 
         Height          =   405
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   258
         Top             =   5730
         Width           =   1455
         Begin VB.OptionButton optDisc 
            Alignment       =   1  'Right Justify
            Caption         =   "-"
            Height          =   225
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   260
            Top             =   150
            Width           =   525
         End
         Begin VB.OptionButton optAdd 
            Alignment       =   1  'Right Justify
            Caption         =   "+"
            Height          =   195
            Left            =   810
            RightToLeft     =   -1  'True
            TabIndex        =   259
            Top             =   150
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin VB.TextBox TxtOfficeValueNet 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   6660
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   256
         Top             =   6090
         Width           =   1245
      End
      Begin VB.TextBox TxtOfficeValueDiscAdd 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   9030
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   254
         Top             =   6120
         Width           =   1395
      End
      Begin VB.TextBox TxtOfficeValue 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   11490
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   252
         Top             =   6120
         Width           =   1245
      End
      Begin VB.TextBox TxtPreBalaValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   246
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox TxtPreBalaPayed 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   245
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox TxtPreBalaRemain 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   244
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox TxtPreBalaTransPyed 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3360
         TabIndex        =   243
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox TxtPreBalaNet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   242
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox TotalPayments 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   11490
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   241
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox TxtNetValue 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   240
         Top             =   6480
         Width           =   11895
      End
      Begin VB.TextBox TxtValuExpenses 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   238
         Top             =   6120
         Width           =   1515
      End
      Begin VB.TextBox TxtPercent 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   236
         Top             =   5760
         Width           =   1515
      End
      Begin VB.TextBox TxtTotalPayedOpBalance 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   3180
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   235
         Top             =   5760
         Width           =   1845
      End
      Begin VB.TextBox TxtNetPayments 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   3180
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   234
         Top             =   6120
         Width           =   1845
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   12960
         RightToLeft     =   -1  'True
         TabIndex        =   217
         Top             =   3480
         Width           =   1200
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H8000000B&
         Caption         =   "«·€«¡ «·”œ«œ"
         Height          =   315
         Left            =   11160
         RightToLeft     =   -1  'True
         TabIndex        =   212
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   12840
         RightToLeft     =   -1  'True
         TabIndex        =   211
         Top             =   900
         Width           =   1200
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
         Height          =   2220
         Left            =   0
         TabIndex        =   213
         Top             =   1200
         Width           =   14280
         _cx             =   25188
         _cy             =   3916
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
         Cols            =   24
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPayments2.frx":038A
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
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
         Height          =   1860
         Left            =   0
         TabIndex        =   216
         Top             =   3840
         Width           =   14280
         _cx             =   25188
         _cy             =   3281
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
         Cols            =   22
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPayments2.frx":0720
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
         Caption         =   "’«›Ì «·« ⁄«»"
         Height          =   255
         Index           =   11
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   257
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6150
         Width           =   1095
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·“Ì«œ… «Ê «·‰ﬁ’«‰"
         Height          =   375
         Index           =   10
         Left            =   10470
         RightToLeft     =   -1  'True
         TabIndex        =   255
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6060
         Width           =   1035
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "« ⁄«» «·„ﬂ » „⁄ ﬁ „"
         Height          =   255
         Index           =   8
         Left            =   12600
         RightToLeft     =   -1  'True
         TabIndex        =   253
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁÌ„…"
         Height          =   255
         Left            =   11400
         RightToLeft     =   -1  'True
         TabIndex        =   251
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„”œœ „”»ﬁ«"
         Height          =   255
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   250
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„ »ﬁÌ"
         Height          =   255
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   249
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„”œœ «·Õ—ﬂ…"
         Height          =   255
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   248
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·’«›Ì «·„” Õﬁ"
         Height          =   255
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   247
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«Ã„«·Ì"
         Height          =   255
         Index           =   4
         Left            =   12480
         RightToLeft     =   -1  'True
         TabIndex        =   239
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "’«›Ì «·œ›⁄« "
         Height          =   255
         Index           =   3
         Left            =   5010
         RightToLeft     =   -1  'True
         TabIndex        =   237
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì ”«»ﬁ «›  «ÕÌ"
         Height          =   255
         Index           =   9
         Left            =   5250
         RightToLeft     =   -1  'True
         TabIndex        =   233
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄„Ê·…"
         Height          =   255
         Index           =   7
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   220
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "»Ì«‰«  „’—Ê›«  «·⁄ﬁ«—"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   219
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   3480
         Width           =   3375
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «·„’—Ê›« "
         Height          =   255
         Index           =   5
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   218
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         Left            =   13920
         RightToLeft     =   -1  'True
         TabIndex        =   215
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «·œ›⁄«  "
         Height          =   255
         Index           =   2
         Left            =   12600
         RightToLeft     =   -1  'True
         TabIndex        =   214
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   5760
         Width           =   1575
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E2E9E9&
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   221
      Top             =   960
      Width           =   6855
      Begin VB.TextBox TxtSearch2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4830
         RightToLeft     =   -1  'True
         TabIndex        =   222
         Top             =   120
         Width           =   825
      End
      Begin MSDataListLib.DataCombo DcbIqara2 
         Height          =   315
         Left            =   180
         TabIndex        =   223
         Top             =   120
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄ﬁ«—"
         Height          =   285
         Index           =   48
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   224
         Top             =   120
         Width           =   1245
      End
   End
   Begin VB.CommandButton CMDSENDSMS 
      Caption         =   "«—”«· —”«·Â"
      Height          =   495
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   207
      Top             =   9240
      Width           =   1095
   End
   Begin VB.TextBox TxtValueTemp 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   206
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H0080FFFF&
      Caption         =   "»Ì«‰«  œ›⁄«  «·„·«ﬂ"
      Height          =   6375
      Index           =   0
      Left            =   -2250
      RightToLeft     =   -1  'True
      TabIndex        =   199
      Top             =   780
      Visible         =   0   'False
      Width           =   16095
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   13800
         RightToLeft     =   -1  'True
         TabIndex        =   201
         Top             =   300
         Width           =   1200
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H8000000B&
         Caption         =   "«·€«¡ «·”œ«œ"
         Height          =   315
         Index           =   0
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   200
         Top             =   240
         Width           =   1695
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   5220
         Left            =   1800
         TabIndex        =   202
         Top             =   600
         Width           =   14280
         _cx             =   25188
         _cy             =   9208
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
         Cols            =   21
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPayments2.frx":0A6E
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
         Caption         =   "≈Ã„«·Ì «·œ›⁄« "
         Height          =   255
         Index           =   0
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   205
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   204
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   5880
         Width           =   8775
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
         Left            =   15600
         RightToLeft     =   -1  'True
         TabIndex        =   203
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·› —Â"
      Height          =   855
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   192
      Top             =   3840
      Width           =   3135
      Begin Dynamic_Byte.NourHijriCal FrmPriodDateH 
         Height          =   315
         Left            =   120
         TabIndex        =   193
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
      End
      Begin MSComCtl2.DTPicker FrmPriodDate 
         Height          =   315
         Left            =   1470
         TabIndex        =   194
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   507379713
         CurrentDate     =   41640
      End
      Begin Dynamic_Byte.NourHijriCal ToPriodDateH 
         Height          =   315
         Left            =   120
         TabIndex        =   195
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
      End
      Begin MSComCtl2.DTPicker ToPriodDate 
         Height          =   315
         Left            =   1470
         TabIndex        =   196
         Top             =   480
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   507379713
         CurrentDate     =   41640
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   63
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   198
         Top             =   120
         Width           =   405
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ï"
         Height          =   285
         Index           =   64
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   197
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1935
      Left            =   14520
      RightToLeft     =   -1  'True
      TabIndex        =   178
      Top             =   1680
      Width           =   6255
      Begin VB.TextBox TxtDiff 
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
         Left            =   360
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   191
         Top             =   600
         Width           =   1425
      End
      Begin VB.TextBox txtComisinold 
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
         Left            =   3360
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   188
         Top             =   600
         Width           =   2025
      End
      Begin VB.TextBox txtinstrancold 
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
         Left            =   3360
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   187
         Top             =   960
         Width           =   2025
      End
      Begin VB.TextBox txtComisin 
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
         TabIndex        =   182
         Top             =   1200
         Width           =   2025
      End
      Begin VB.TextBox txttotal1 
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
         Left            =   3360
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   181
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox txtinstranc 
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
         TabIndex        =   180
         Top             =   1560
         Width           =   2025
      End
      Begin VB.TextBox txttotal2 
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
         Left            =   360
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   179
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·”⁄Ì"
         Height          =   195
         Index           =   11
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   190
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·«ÌÃ«—"
         Height          =   195
         Index           =   1
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   189
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·”⁄Ì"
         Height          =   195
         Index           =   8
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   186
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì1"
         Height          =   195
         Index           =   9
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   185
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·«ÌÃ«—"
         Height          =   195
         Index           =   10
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   184
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì 2"
         Height          =   195
         Index           =   7
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   183
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2E9E9&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   160
      Top             =   960
      Width           =   6855
      Begin VB.TextBox TxtNotID 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   165
         Top             =   120
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox TxtNotVal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   162
         Top             =   120
         Width           =   1515
      End
      Begin VB.TextBox TxtNotSreail1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   161
         Top             =   150
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «Ã„«·Ì "
         Height          =   285
         Index           =   47
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   164
         Top             =   120
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ ”‰œ «·⁄—»Ê‰"
         Height          =   285
         Index           =   46
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   163
         Top             =   120
         Width           =   1275
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2E9E9&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   145
      Top             =   960
      Width           =   6855
      Begin VB.TextBox TxtFilterNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   148
         Top             =   150
         Width           =   1515
      End
      Begin VB.TextBox TXtFilter 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   147
         Top             =   120
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «· ’›ÌÂ"
         Height          =   285
         Index           =   60
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   150
         Top             =   120
         Width           =   795
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «Ã„«·Ì "
         Height          =   285
         Index           =   61
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   149
         Top             =   120
         Width           =   555
      End
   End
   Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
      Height          =   315
      Left            =   7680
      TabIndex        =   142
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
   End
   Begin VB.Frame Frame4 
      Caption         =   "„’«—Ì› ÕÊ«·… »‰ﬂÌ…"
      Height          =   615
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   138
      Top             =   3120
      Width           =   2895
      Begin VB.TextBox txtTransferExpenses 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   240
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "—”Ê„ «·ÕÊ«·Â"
         Height          =   255
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   139
         ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
         Top             =   240
         Width           =   975
      End
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
      Left            =   7200
      TabIndex        =   137
      Top             =   1995
      Width           =   2805
   End
   Begin VB.TextBox XPTxtID1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15720
      RightToLeft     =   -1  'True
      TabIndex        =   133
      Text            =   "Text4"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
      Height          =   3195
      Index           =   2
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   121
      Top             =   4680
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox TxtPaymentCounts 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   5310
         MaxLength       =   2
         TabIndex        =   126
         Top             =   240
         Width           =   825
      End
      Begin VB.ComboBox CmbMonth 
         Height          =   315
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   125
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox ChkSaleryDis 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈Œ’„ „‰ «·„— »  ·ﬁ«∆Ì«"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3720
         TabIndex        =   124
         Top             =   240
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.ComboBox CboYear 
         Height          =   315
         Left            =   3750
         Style           =   2  'Dropdown List
         TabIndex        =   123
         Top             =   960
         Width           =   1095
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   435
         Index           =   11
         Left            =   4590
         TabIndex        =   122
         Top             =   1320
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
         ButtonImage     =   "FrmPayments2.frx":0DAD
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2565
         Left            =   90
         TabIndex        =   127
         Top             =   210
         Width           =   3495
         _cx             =   6165
         _cy             =   4524
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
         FormatString    =   $"FrmPayments2.frx":1147
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
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   141
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·œ›⁄« "
         Height          =   285
         Index           =   44
         Left            =   6030
         TabIndex        =   132
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «Ê· œ›⁄…"
         Height          =   285
         Index           =   43
         Left            =   5460
         TabIndex        =   131
         Top             =   690
         Width           =   1305
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
         Height          =   495
         Index           =   0
         Left            =   4500
         TabIndex        =   130
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Â—"
         Height          =   315
         Index           =   42
         Left            =   4890
         TabIndex        =   129
         Top             =   630
         Width           =   405
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”‰…"
         Height          =   315
         Index           =   41
         Left            =   4890
         TabIndex        =   128
         Top             =   960
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„«  «·ÕÊ«·Â"
      Height          =   1815
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   115
      Top             =   11280
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   116
         Top             =   240
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   120
         TabIndex        =   117
         Top             =   570
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Format          =   507379713
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
         TabIndex        =   119
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
         TabIndex        =   118
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox TxtCustCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12210
      RightToLeft     =   -1  'True
      TabIndex        =   114
      Top             =   1560
      Width           =   705
   End
   Begin VB.TextBox txt_ORDER_NO 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3900
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txt_general_des 
      Alignment       =   1  'Right Justify
      Height          =   1605
      Left            =   7170
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5280
      Width           =   5775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì Õ«·… «·„ÊŸ›"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   109
      Top             =   960
      Width           =   6855
      Begin VB.OptionButton Option7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»œ·«  „ﬁœ„Â"
         Height          =   195
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   140
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Œ’’« "
         Height          =   195
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   136
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«ÃÊ— „” Õﬁ…"
         Height          =   195
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   111
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”·›…"
         Height          =   195
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   110
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtNoteSerial1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   11475
      RightToLeft     =   -1  'True
      TabIndex        =   106
      Top             =   600
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   104
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAdv_payment_value 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   3870
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   101
      Top             =   1515
      Width           =   1965
   End
   Begin VB.Frame Frame1 
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
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   95
      Top             =   480
      Width           =   3735
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
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   98
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
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   97
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
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   720
         Width           =   2055
      End
      Begin ALLButtonS.ALLButton ALLButton3 
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   " ÕœÌœ «·œ›⁄« "
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
         MICON           =   "FrmPayments2.frx":11D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton ALLButton4 
         Height          =   255
         Left            =   120
         TabIndex        =   209
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   " ›«’Ì· «·⁄ﬁ«—"
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
         MICON           =   "FrmPayments2.frx":11EE
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
      Left            =   -90
      RightToLeft     =   -1  'True
      TabIndex        =   79
      Top             =   7920
      Width           =   11295
      Begin VB.TextBox Txtownerid 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   208
         Top             =   240
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8520
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   107
         Top             =   200
         Width           =   1785
      End
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   80
         Top             =   180
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboCreditSide 
         Height          =   315
         Left            =   90
         TabIndex        =   81
         Top             =   510
         Width           =   5295
         _ExtentX        =   9340
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
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   180
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› œ«∆‰"
         Height          =   285
         Index           =   31
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   510
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   315
         Index           =   30
         Left            =   10170
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   210
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·› —… :"
         Height          =   315
         Index           =   29
         Left            =   10170
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   540
         Width           =   975
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   83
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
         TabIndex        =   82
         Top             =   510
         Width           =   1485
      End
   End
   Begin VB.Frame FraNote 
      BackColor       =   &H00E2E9E9&
      Height          =   1365
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   2400
      Width           =   7155
      Begin VB.TextBox txtperson 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   5685
      End
      Begin VB.TextBox TxtChequeNumber 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   570
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker DtpChequeDueDate 
         Height          =   315
         Left            =   30
         TabIndex        =   9
         Top             =   570
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Format          =   200474625
         CurrentDate     =   39614
      End
      Begin MSDataListLib.DataCombo DcboBankName 
         Height          =   315
         Left            =   30
         TabIndex        =   7
         Top             =   150
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboBox 
         Height          =   315
         Left            =   3510
         TabIndex        =   6
         Top             =   150
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  .«·≈” Õﬁ«ﬁ"
         Height          =   285
         Index           =   17
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   232
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·‘Ìﬂ"
         Height          =   285
         Index           =   16
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   231
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„” ›Ìœ"
         Height          =   285
         Index           =   34
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·Œ“Ì‰Â"
         Height          =   285
         Index           =   9
         Left            =   5790
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·»‰ﬂ"
         Height          =   285
         Index           =   15
         Left            =   1950
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   150
         Width           =   1215
      End
   End
   Begin VB.ComboBox CboPaymentType 
      Height          =   315
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1995
      Width           =   1935
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
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   2490
      Width           =   3705
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   0
         Left            =   1830
         TabIndex        =   56
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
         MouseIcon       =   "FrmPayments2.frx":120A
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
         TabIndex        =   57
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
         MouseIcon       =   "FrmPayments2.frx":136C
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
         TabIndex        =   58
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
         MouseIcon       =   "FrmPayments2.frx":14CE
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
         TabIndex        =   59
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
         MouseIcon       =   "FrmPayments2.frx":1630
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
         TabIndex        =   60
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
         MouseIcon       =   "FrmPayments2.frx":1792
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
         TabIndex        =   61
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
         MouseIcon       =   "FrmPayments2.frx":18F4
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
         TabIndex        =   62
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
         MouseIcon       =   "FrmPayments2.frx":1A56
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
         TabIndex        =   63
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
         MouseIcon       =   "FrmPayments2.frx":1BB8
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
         TabIndex        =   64
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
         MouseIcon       =   "FrmPayments2.frx":1D1A
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
         Top             =   780
         Width           =   675
      End
   End
   Begin VB.CheckBox ChkTrans 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰ Õ”«» ›« Ê—…"
      Height          =   225
      Left            =   17610
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   390
      Width           =   1575
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   975
      Index           =   0
      Left            =   18990
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   630
      Width           =   3675
      Begin VB.TextBox TxtTransID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   180
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TxtTransSerial 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   540
         Width           =   1005
      End
      Begin VB.ComboBox CboTrans 
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   210
         Width           =   1995
      End
      Begin ImpulseButton.ISButton CmdSearchTrans 
         Height          =   345
         Left            =   600
         TabIndex        =   37
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
         ButtonImage     =   "FrmPayments2.frx":1E7C
      End
      Begin ImpulseButton.ISButton CmdOpenTrans 
         Height          =   345
         Left            =   90
         TabIndex        =   39
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
         ButtonImage     =   "FrmPayments2.frx":2216
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   270
         Width           =   1305
      End
   End
   Begin VB.ComboBox DCboCashType 
      Height          =   315
      ItemData        =   "FrmPayments2.frx":25B0
      Left            =   11460
      List            =   "FrmPayments2.frx":25B2
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   945
      Width           =   1455
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   1365
      Left            =   7170
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3900
      Width           =   5775
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10260
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1995
      Width           =   2655
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   585
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   14385
      _cx             =   25374
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
      Caption         =   " ”‰œ ’—› - «·„œ›Ê⁄«   ··«„·«ﬂ   "
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
      Begin VB.TextBox oldtxtNoteSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   134
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
         TabIndex        =   89
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
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1155
         TabIndex        =   17
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
         ButtonImage     =   "FrmPayments2.frx":25B4
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
         TabIndex        =   18
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
         ButtonImage     =   "FrmPayments2.frx":294E
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
         TabIndex        =   19
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
         ButtonImage     =   "FrmPayments2.frx":2CE8
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
         TabIndex        =   20
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
         ButtonImage     =   "FrmPayments2.frx":3082
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
         Left            =   2040
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
         Left            =   0
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
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   6360
         Picture         =   "FrmPayments2.frx":341C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   9210
      TabIndex        =   0
      Top             =   600
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      Format          =   507445249
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   8700
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   6270
      TabIndex        =   27
      Top             =   8820
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   13440
      TabIndex        =   42
      Top             =   9330
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
      Left            =   12585
      TabIndex        =   43
      Top             =   9330
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
      Left            =   11760
      TabIndex        =   44
      Top             =   9330
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
      Left            =   10815
      TabIndex        =   45
      Top             =   9330
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
      Left            =   9930
      TabIndex        =   46
      Top             =   9330
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
      Left            =   1110
      TabIndex        =   47
      Top             =   9330
      Width           =   615
      _ExtentX        =   1085
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
      Left            =   1755
      TabIndex        =   48
      Top             =   9330
      Width           =   735
      _ExtentX        =   1296
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
      Left            =   9045
      TabIndex        =   49
      Top             =   9330
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
      Left            =   8190
      TabIndex        =   50
      Top             =   9330
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
      TabIndex        =   52
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
      MouseIcon       =   "FrmPayments2.frx":7084
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
      TabIndex        =   53
      Top             =   9480
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
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   15960
      TabIndex        =   90
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«ŸÂ«— «·«ﬁ”«ÿ"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPayments2.frx":71E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Left            =   16080
      TabIndex        =   91
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«ŸÂ«— ”‰œ «·„œÌÊ‰Ì…"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPayments2.frx":7202
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DCPROJECT 
      Height          =   315
      Left            =   14640
      TabIndex        =   92
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
      Bindings        =   "FrmPayments2.frx":721E
      Height          =   315
      Left            =   3900
      TabIndex        =   14
      Top             =   2760
      Width           =   1935
      _ExtentX        =   3413
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
      Left            =   2640
      TabIndex        =   105
      Top             =   9330
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
      Left            =   4800
      TabIndex        =   108
      Top             =   9720
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
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   3600
      TabIndex        =   143
      Top             =   9330
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
      Left            =   7080
      TabIndex        =   144
      Top             =   9330
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
   Begin ImpulseButton.ISButton ISButton3 
      Height          =   375
      Left            =   10920
      TabIndex        =   146
      TabStop         =   0   'False
      ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
      Top             =   840
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
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
      ButtonImage     =   "FrmPayments2.frx":7233
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic2 
      Height          =   855
      Left            =   7200
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   6960
      Width           =   7095
      _cx             =   12515
      _cy             =   1508
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
      Begin VB.TextBox TxtSearch 
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
         Left            =   4680
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   152
         Top             =   120
         Width           =   1035
      End
      Begin MSDataListLib.DataCombo DcbIqara 
         Height          =   315
         Left            =   120
         TabIndex        =   153
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
         Top             =   120
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbUnitNo 
         Height          =   315
         Left            =   120
         TabIndex        =   154
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbUnitType 
         Height          =   315
         Left            =   3540
         TabIndex        =   155
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ÊÕœ…"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   14
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   158
         Top             =   480
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·ÊÕœ…"
         Height          =   195
         Index           =   15
         Left            =   5985
         RightToLeft     =   -1  'True
         TabIndex        =   157
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄ﬁ«—"
         Height          =   195
         Index           =   4
         Left            =   5985
         RightToLeft     =   -1  'True
         TabIndex        =   156
         Top             =   120
         Width           =   990
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   375
      Left            =   10920
      TabIndex        =   159
      TabStop         =   0   'False
      ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
      Top             =   840
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   4
      Caption         =   ""
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
      ButtonImage     =   "FrmPayments2.frx":7630
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic3 
      Height          =   3375
      Left            =   14880
      TabIndex        =   166
      TabStop         =   0   'False
      Top             =   3000
      Width           =   6375
      _cx             =   11245
      _cy             =   5953
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
      BackColor       =   -2147483633
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
      Begin VB.TextBox TxtRent 
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
         TabIndex        =   172
         Top             =   1440
         Width           =   2025
      End
      Begin VB.TextBox txtWater 
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
         Left            =   3120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   171
         Top             =   2160
         Width           =   2025
      End
      Begin VB.TextBox txtinstrunce 
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
         TabIndex        =   170
         Top             =   2160
         Width           =   2025
      End
      Begin VB.TextBox TxtCommission 
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
         Left            =   3120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   169
         Top             =   1800
         Width           =   2025
      End
      Begin VB.Frame Frame7 
         Height          =   615
         Left            =   120
         TabIndex        =   168
         Top             =   2640
         Width           =   5055
      End
      Begin VB.TextBox TxtCommissionOut 
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
         TabIndex        =   167
         Top             =   1800
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·”⁄Ì"
         Height          =   195
         Index           =   2
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   177
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·«ÌÃ«—"
         Height          =   195
         Index           =   3
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   176
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·„Ì«Â"
         Height          =   195
         Index           =   5
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   175
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «· «„Ì‰"
         Height          =   195
         Index           =   6
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   174
         Top             =   2160
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "”⁄Ì „ﬂ » Œ«—ÃÌ"
         Height          =   195
         Index           =   12
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   173
         Top             =   1800
         Width           =   990
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   13
      Left            =   4560
      TabIndex        =   225
      Top             =   9330
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… 2"
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
      Left            =   5400
      TabIndex        =   226
      Top             =   9330
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… ⁄„Ê·«  «·«„·«ﬂ"
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
   Begin VB.TextBox TxtTotalInsurances 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   7200
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   227
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «„Ì‰"
      Height          =   285
      Index           =   49
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   228
      Top             =   1560
      Width           =   405
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
      TabIndex        =   135
      Top             =   9840
      Width           =   3675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   195
      Index           =   40
      Left            =   6780
      RightToLeft     =   -1  'True
      TabIndex        =   120
      Top             =   600
      Width           =   765
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»Ì…"
      Height          =   285
      Index           =   37
      Left            =   5850
      RightToLeft     =   -1  'True
      TabIndex        =   113
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·‘—Õ «·⁄«„"
      Height          =   285
      Index           =   36
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   112
      Top             =   5880
      Width           =   1155
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
      Height          =   255
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   103
      Top             =   2760
      Width           =   1215
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
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   102
      Top             =   1575
      Width           =   1245
   End
   Begin VB.Label lblsqlstring 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   135
      Left            =   16800
      RightToLeft     =   -1  'True
      TabIndex        =   100
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„‘—Ê⁄"
      Height          =   285
      Index           =   33
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   93
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
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   1980
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
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
      Height          =   405
      Index           =   18
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   2010
      Width           =   3555
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—’Ìœ «·Õ«·Ï:"
      Height          =   285
      Index           =   13
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   1650
      Width           =   1185
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Left            =   3270
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   8850
      Width           =   825
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Left            =   1500
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   8850
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   255
      Index           =   6
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   8850
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   255
      Index           =   7
      Left            =   4140
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   8850
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   300
      Index           =   8
      Left            =   8220
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   8850
      Width           =   1140
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„œ›Ê⁄« "
      Height          =   285
      Index           =   0
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   10680
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   615
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… «·„œ›Ê⁄« "
      Height          =   285
      Index           =   2
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2010
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„"
      Height          =   285
      Index           =   3
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1545
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·”‰œ"
      Height          =   285
      Index           =   4
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   630
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "»‰«¡ ⁄·Ï"
      Height          =   285
      Index           =   5
      Left            =   12630
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4770
      Width           =   1455
   End
End
Attribute VB_Name = "FrmPayments2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim Line1 As Double
Dim Line2 As Double
Dim Line3 As Double
Dim departement_name As Integer
Dim numbering_type As Integer
Dim Balance As String
Dim balanceString As String
Dim Account_Code_dynamic As String
Public called As Boolean
Dim FlgBillBuy As Boolean
 Function OtherOwnerNoreatJlInContractFiter(LngDevID As Long, notes_id As Double) As Double

If DCboCashType.ListIndex <> 6 Then Exit Function
Dim Percetage As Double
Dim commissionvalue As Double
Dim total_value As Double
Dim cProgress As ClsProgress
Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double
 Dim lineno As Double

         lineno = 1
 Dim AccountCode As String
    cProgress.StartProgress
    DoEvents
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim Msgdes As String
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
     Dim i As Integer
     Dim AccountCodeDept As String
Msgdes = "»‰«¡ ⁄·Ï „œ›Ê⁄«   ’›Ì… «„·«ﬂ «·€Ì— —ﬁ„ " & TxtNoteSerial1.text & " "
Dim AccountCodeVat As String
Msg = XPMTxtRemarks.text & CHR(13) & Msgdes
AccountCodeDept = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
                  If SystemOptions.OpenAccountAqar = False Then
                        AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                    Else
                        AccountCode = GetAqarAcountCode(val(DcbIqara.BoundText))
                    End If
'AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
                total_value = GetValueFiter(val(TxtFilterNo.text), "RemainRent") 'val(XPTxtVal.Text)
                If ComResid(1).value = True Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                commissionvalue = total_value * Percetage / 100
              commissionvalue = Round(commissionvalue, 2)
               Else
               commissionvalue = 0
               End If
               
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«·«ÌÃ«— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                   
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  «·«ÌÃ«—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…··«ÌÃ«— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
           ''//////—’Ìœ ”«»ﬁ
                           total_value = GetValueFiter(val(TxtFilterNo.text), "OldRent") 'val(XPTxtVal.Text)
                If ComResid(1).value = True Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                commissionvalue = total_value * Percetage / 100
              commissionvalue = Round(commissionvalue, 2)
               Else
               commissionvalue = 0
               End If
               
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "—’Ìœ ”«»ﬁ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                   
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  —’Ìœ ”«»ﬁ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…—’Ìœ ”«»ﬁ   ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              ''//////////«ÌÃ«— «Ì«„ “Ì«œ…
                    total_value = GetValueFiterHeader(val(TxtFilterNo.text), "DaysValueIncrease") 'val(XPTxtVal.Text)
                If ComResid(1).value = True Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                commissionvalue = total_value * Percetage / 100
              commissionvalue = Round(commissionvalue, 2)
               Else
               commissionvalue = 0
               End If
               
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«·«ÌÃ«— «Ì«„ “Ì«œ…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                   
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  «·«ÌÃ«— «Ì«„ “Ì«œ…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…··«ÌÃ«— «·«ÌÃ«— «Ì«„ “Ì«œ…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
                ''//////„Ì«Â
                          total_value = GetValueFiter(val(TxtFilterNo.text), "RemainWater") 'val(XPTxtVal.Text)
               If total_value > 0 Then
                If ComResid(1).value = True Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
                 Else
                     commissionvalue = 0
                End If
             
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "„Ì«Â ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " „Ì«Â ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…··„Ì«Â ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              ''//////ﬂÂ—»«¡
              
                     total_value = GetValueFiter(val(TxtFilterNo.text), "BillPrice") 'val(XPTxtVal.Text)
                  If total_value > 0 Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "ﬂÂ—»«¡ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  ﬂÂ—»«¡", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…··ﬂÂ—»«¡ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
                         ''//////Œœ„« 
              
                     total_value = GetValueFiter(val(TxtFilterNo.text), "RemainService") 'val(XPTxtVal.Text)
                  If total_value > 0 Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
     
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & " Œœ„« ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   If Percetage > 0 And ComResid(1).value = True Then
                    total_value = total_value / (Percetage / 100 + 1)
                    commissionvalue = total_value * Percetage / 100
                    Else
                    commissionvalue = 0
                    End If
                 
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  Œœ„« ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…«·Œœ„«  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
              
                                     ''//////”⁄Ì
              
                     total_value = GetValueFiter(val(TxtFilterNo.text), "RemainCommissions") 'val(XPTxtVal.Text)
                  If total_value > 0 Then
                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & " ”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   If Percetage > 0 And ComResid(1).value = True Then
                    total_value = total_value / (Percetage / 100 + 1)
                    commissionvalue = total_value * Percetage / 100
                    Else
                    commissionvalue = 0
                    End If
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  ”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›…··”⁄Ì ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
   
                            
                         ''''// «· «„Ì‰
              
            total_value = val(txtTotalinsuranceS.text) - GetValueFiter(val(TxtFilterNo.text), "insurance")
            total_value = Abs(total_value)
             If total_value > 0 Then
                     
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & " «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              
                                       ''''// «· «„Ì‰
              
            total_value = GetValueFiter(val(TxtFilterNo.text), "insurance")
            total_value = Abs(total_value)
             If total_value > 0 Then
                     
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & " «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                      
                             If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 0, Msg & " " & " «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                 
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & " «„Ì‰", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                      
              End If
              
              
             

              
                           total_value = val(XPTxtVal.text)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, DcboDebitSide.BoundText, total_value, 0, Msg & " " & "«·’‰œÊﬁ/«·»‰ﬂ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, DcboCreditSide.BoundText, total_value, 1, Msg & " " & "«·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              

        
    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
ErrTrap:
End Function
' Function OtherOwnerNoreatJlInContractFiter111(LngDevID As Long, notes_id As Double) As Double
'
'If DCboCashType.ListIndex <> 6 Then Exit Function
'Dim Percetage As Double
'Dim commissionvalue As Double
'Dim total_value As Double
'Dim cProgress As ClsProgress
'Set cProgress = New ClsProgress
'    cProgress.ProgressType = Waiting
' Dim foxy_ked_NO As String
' Dim credit_side As String
' Dim My_SQL As String
' Dim Line1 As Double
' Dim lineno As Double
'         lineno = 1
' Dim AccountCode As String
'    cProgress.StartProgress
'    DoEvents
'    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'    Dim Msgdes As String
'    Dim CURRENT_LINE As Double
'    Dim depit_side As String
'    Dim Msg As String
'     Dim i As Integer
'Msgdes = "»‰«¡ ⁄·Ï „ﬁ»Ê÷«   ’›Ì… «„·«ﬂ «·€Ì— —ﬁ„ " & TxtNoteSerial1.Text & " "
'Dim AccountCodeVat As String
'Msg = XPMTxtRemarks.Text & Chr(13) & Msgdes
'                total_value = val(TxtTotalInsurances.Text) + val(XPTxtVal.Text)
'                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
'                commissionvalue = total_value * Percetage / 100
'              commissionvalue = Round(commissionvalue, 2)
'
'             If total_value > 0 Then
'             AccountCode = get_account_code_branch(86, my_branch)
'                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value - commissionvalue, 0, Msg & " " & "«·«Ì—«œ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                    lineno = lineno + 1
'                     If commissionvalue > 0 Then
'
'                If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 0, Msg & " " & " «·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                      lineno = lineno + 1
'                    End If
'
'                    AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
'                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  «·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                      lineno = lineno + 1
'
'              End If
'              ''''// «· «„Ì‰
'
'            total_value = val(TxtTotalInsurances.Text)
'             If total_value > 0 Then
'                     AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
'                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 0, Msg & " " & "«·„«·ﬂ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                    lineno = lineno + 1
'                     AccountCode = get_account_code_branch(81, my_branch)
'                         If ModAccounts.AddNewDev(LngDevID, lineno, DcboCreditSide.BoundText, total_value, 1, Msg & " " & "«·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                      lineno = lineno + 1
'              End If
'
'
'
'                total_value = val(XPTxtVal.Text) + commissionvalue
'             If total_value > 0 Then
'                    If ModAccounts.AddNewDev(LngDevID, lineno, DcboDebitSide.BoundText, total_value, 0, Msg & " " & "«·’‰œÊﬁ/«·»‰ﬂ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                    lineno = lineno + 1
'
'                         If ModAccounts.AddNewDev(LngDevID, lineno, DcboCreditSide.BoundText, total_value, 1, Msg & " " & "«·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
''                      lineno = lineno + 1
 '             End If
 '
'
'
'    DoEvents
'    cProgress.FinishProgress
'    cProgress.StopProgess
'    Set cProgress = Nothing
'
'ErrTrap:
'End Function
 Function MyOwnerNoreatJlInContractFiter(LngDevID As Long, notes_id As Double) As Double

If DCboCashType.ListIndex <> 6 Then Exit Function
Dim Percetage As Double
Dim commissionvalue As Double
Dim total_value As Double
Dim cProgress As ClsProgress
Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
 Dim foxy_ked_NO As String
 Dim credit_side As String
 Dim My_SQL As String
 Dim Line1 As Double
 Dim lineno As Double
 Dim AccountCodeDept As String
         lineno = 1
 Dim AccountCode As String
    cProgress.StartProgress
    DoEvents
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim Msgdes As String
    Dim CURRENT_LINE As Double
    Dim depit_side As String
    Dim Msg As String
     Dim i As Integer
     
     Dim AccountCodeVat As String
Msgdes = "»‰«¡ ⁄·Ï „œ›Ê⁄«   ’›Ì… «„·«ﬂÌ  —ﬁ„ " & TxtNoteSerial1.text & " "
PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage

Msg = XPMTxtRemarks.text & CHR(13) & Msgdes
 AccountCodeDept = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
                total_value = GetValueFiter(val(TxtFilterNo.text), "RemainRent")
                If total_value > 0 Then
              If ComResid(1).value = True Then
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«·„” «Ã— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    AccountCode = get_account_code_branch(86, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «Ì—«œ«  «·«ÌÃ«—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
           '////—’Ìœ ”«»ﬁ
                          total_value = GetValueFiter(val(TxtFilterNo.text), "OldRent")
                If total_value > 0 Then
              If ComResid(1).value = True Then
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "—’Ìœ ”«»ﬁ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    AccountCode = get_account_code_branch(86, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  —’Ìœ ”«»ﬁ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… —’Ìœ ”«»ﬁ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              '////«·«ÌÃ«— «Ì«„ “Ì«œ…
                              total_value = GetValueFiterHeader(val(TxtFilterNo.text), "DaysValueIncrease")
                If total_value > 0 Then
              If ComResid(1).value = True Then
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«ÌÃ«— «Ì«„ “Ì«œ… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    AccountCode = get_account_code_branch(86, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  «Ã«— «Ì„ “Ì«œ…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… «ÌÃ«— «Ì«„ “Ì«œ…", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
              ''////„Ì«Â
                              total_value = GetValueFiter(val(TxtFilterNo.text), "RemainWater")
                If total_value > 0 Then

             If ComResid(1).value = True Then
                     
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«·„” «Ã— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1

                    AccountCode = get_account_code_branch(83, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «Ì—«œ«  «·„Ì«Â", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… ··„Ì«Â ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
              ''////«·ﬂÂ—»«¡
                  total_value = GetValueFiter(val(TxtFilterNo.text), "BillPrice")
                If total_value > 0 Then
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)

                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value + commissionvalue, 0, Msg & " " & "«·„” «Ã— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                   
                    AccountCode = get_account_code_branch(84, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «Ì—«œ«  «·ﬂÂ—»«¡", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… ··ﬂÂ—»«¡ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              
              
                            ''////Œœ„« 
                              total_value = GetValueFiter(val(TxtFilterNo.text), "RemainService")
                If total_value > 0 Then

            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & "«·„” «Ã— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
            If ComResid(1).value = True Then
                     total_value = total_value / (Percetage / 100 + 1)
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
              
                    AccountCode = get_account_code_branch(85, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «Ì—«œ«  Œœ„« ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… Œœ„«  ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
                                          ''////”⁄Ì
                              total_value = GetValueFiter(val(TxtFilterNo.text), "RemainCommissions")
                If total_value > 0 Then
            
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & "«·„” «Ã— ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
              If ComResid(1).value = True Then
                     total_value = total_value / (Percetage / 100 + 1)
                     commissionvalue = total_value * Percetage / 100
                     commissionvalue = Round(commissionvalue, 2)
              Else
              commissionvalue = 0
              End If
                    AccountCode = get_account_code_branch(81, my_branch)
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «Ì—«œ«  «·”⁄Ì", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    If commissionvalue > 0 Then
                    
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 1, Msg & " " & " «·ﬁÌ„… «·„÷«›… ··”⁄Ì ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                    End If
              End If
              ''''// «· «„Ì‰
              
            total_value = val(txtTotalinsuranceS.text) - GetValueFiter(val(TxtFilterNo.text), "insurance")
            total_value = Abs(total_value)
             If total_value > 0 Then
                    AccountCode = get_account_code_branch(82, my_branch)
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
                            
                        ''''// «· «„Ì‰
              
            total_value = GetValueFiter(val(TxtFilterNo.text), "insurance")
             If total_value > 0 Then
                    AccountCode = get_account_code_branch(82, my_branch)
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 0, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                      
                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 0, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
                      
              End If
              
             
                total_value = val(XPTxtVal.text)
             If total_value > 0 Then
                    If ModAccounts.AddNewDev(LngDevID, lineno, DcboDebitSide.BoundText, total_value, 0, Msg & " " & "«·’‰œÊﬁ/«·»‰ﬂ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    lineno = lineno + 1
                     
                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeDept, total_value, 1, Msg & " " & "«·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(dcBranch.BoundText), , , , , , , , , , val(DcbIqara.BoundText), val(DcbUnitType.BoundText), val(DcbUnitNo.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                      lineno = lineno + 1
              End If
              

        
    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
    
ErrTrap:
End Function
' Function MyOwnerNoreatJlInContractFiter1111(LngDevID As Long, notes_id As Double) As Double
'
'If DCboCashType.ListIndex <> 6 Then Exit Function
'Dim Percetage As Double
'Dim commissionvalue As Double
'Dim total_value As Double
'Dim cProgress As ClsProgress
'Set cProgress = New ClsProgress
'    cProgress.ProgressType = Waiting
' Dim foxy_ked_NO As String
' Dim credit_side As String
' Dim My_SQL As String
' Dim Line1 As Double
' Dim lineno As Double
'         lineno = 1
' Dim AccountCode As String
'    cProgress.StartProgress
'    DoEvents
'    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'    Dim Msgdes As String
'    Dim CURRENT_LINE As Double
'    Dim depit_side As String
'    Dim Msg As String
'     Dim i As Integer
'Msgdes = "»‰«¡ ⁄·Ï „ﬁ»Ê÷«   ’›Ì… «„·«ﬂ «·€Ì— —ﬁ„ " & txtNoteSerial1.Text & " "
'Dim AccountCodeVat As String
'Msg = XPMTxtRemarks.Text & Chr(13) & Msgdes
'                total_value = val(TxtTotalInsurances.Text) + val(XPTxtVal.Text)
'                     PercentgValueAddedAccount_Transec XPDtbTrans.value, 8, 1, AccountCodeVat, Percetage
'                commissionvalue = total_value * Percetage / 100
'              commissionvalue = Round(commissionvalue, 2)
'
'             If total_value > 0 Then
'             AccountCode = get_account_code_branch(86, my_branch)
'                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value - commissionvalue, 0, Msg & " " & "«·«Ì—«œ ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(DcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                    lineno = lineno + 1
'                     If commissionvalue > 0 Then
'
'                If ModAccounts.AddNewDev(LngDevID, lineno, AccountCodeVat, commissionvalue, 0, Msg & " " & " «·ﬁÌ„… «·„÷«›… ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(DcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                      lineno = lineno + 1
'                    End If
'
'                    AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
'                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "  «·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(DcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                      lineno = lineno + 1
'
'              End If
'              ''''// «· «„Ì‰
'
'            total_value = val(TxtTotalInsurances.Text)
'             If total_value > 0 Then
'              AccountCode = get_account_code_branch(82, my_branch)
'
'                    If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 0, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(DcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                    lineno = lineno + 1
'                     AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
'                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & " «„Ì‰ „” —œ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(DcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                      lineno = lineno + 1
'              End If
'
'
'
'                total_value = val(XPTxtVal.Text) + commissionvalue
'             If total_value > 0 Then
'                    If ModAccounts.AddNewDev(LngDevID, lineno, DcboCreditSide.BoundText, total_value, 0, Msg & " " & "«·’‰œÊﬁ/«·»‰ﬂ", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(DcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                    lineno = lineno + 1
'                     AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(DBCboClientName.BoundText), "Account_code")
'                         If ModAccounts.AddNewDev(LngDevID, lineno, AccountCode, total_value, 1, Msg & " " & "«·„” «Ã—", val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , val(DcBranch.BoundText)) = False Then
'                        GoTo ErrTrap
'                    End If
'                      lineno = lineno + 1
'              End If
'
'
'
'    DoEvents
'    cProgress.FinishProgress
'    cProgress.StopProgess
'    Set cProgress = Nothing
'
'ErrTrap:
'End Function
Private Sub ALLButton4_Click()

 TotalPayments.text = 0
TxtTotalPayedOpBalance.text = 0
txtPercent.text = 0
TxtValuExpenses.text = 0
TxtNetPayments.text = 0
TxtOfficeValue.text = 0
TxtOfficeValueNet = 0
TxtOfficeValueDiscAdd = 0
TxtNetValue.text = 0

    lblsqlstring.Caption = ""
If val(DBCboClientName.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— «·„«·ﬂ «Ê·«"
Else
MsgBox "Please Select Owner"
End If
'DBCboClientName.SetFocus
Exit Sub
Else
Frame12(1).Visible = True

If Me.TxtModFlg.text <> "R" Then
XPTxtVal.text = 0
If Me.TxtModFlg.text = "N" Then
CalculteOpinigBala
RetriveOwnerPayment202 val(DBCboClientName.BoundText)
RetriveOwnerPayment203 val(DBCboClientName.BoundText)
End If
If Me.TxtModFlg.text = "E" Then
Command10_Click (0)
End If

If Me.TxtModFlg.text = "E" And (FlgBillBuy = True Or VSFlexGrid2.rows = 1) Then

Command10_Click (0)
RetriveOwnerPayment202 val(DBCboClientName.BoundText)
RetriveOwnerPayment203 val(DBCboClientName.BoundText)
CalculteOpinigBala
End If
End If
End If
End Sub
Sub CalculteOpinigBala()
Dim AccountCode As String
AccountCode = GetAqarAcountCode(val(DcbIqara.BoundText))
If AccountCode <> "" Then
TxtPreBalaValue.text = GetBalanceValue(AccountCode, 0) - GetBalanceValue(AccountCode, 1)
Else
TxtPreBalaValue = 0
End If
If TxtPreBalaValue > 0 Then
TxtPreBalaValue = 0
Else
TxtPreBalaValue = Abs(TxtPreBalaValue)

End If
TxtPreBalaPayed.text = GeteOwnerPayedOPening()
TxtPreBalaRemain.text = val(TxtPreBalaValue.text) - val(TxtPreBalaPayed.text)
End Sub
Function GetBalanceValue(Optional AcountCode As String, Optional Credit_Or_Debit As Integer = 0) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "SELECT     SUM([Value]) AS Value"
sql = sql & " From dbo.DOUBLE_ENTREY_VOUCHERS1"
sql = sql & " WHERE     (Account_Code = N'" & AcountCode & "') AND (Credit_Or_Debit = " & Credit_Or_Debit & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetBalanceValue = IIf(IsNull(rs2("Value").value), 0, rs2("Value").value)
Else
GetBalanceValue = 0
End If
End Function

Private Sub C1Elastic1_Click()
FrmPayments.show
End Sub

Private Sub Check1_Click()
    Dim i As Integer

    If Check1.value = vbChecked Then

        With Me.VSFlexGrid1
 
            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
                .TextMatrix(i, .ColIndex("TransPayedValue")) = .TextMatrix(i, .ColIndex("RemainingValue"))
            Next i

        End With

    Else

        With Me.VSFlexGrid1

            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("payed")) = False
                .TextMatrix(i, .ColIndex("TransPayedValue")) = 0
            Next i

        End With

    End If
End Sub

Private Sub Check2_Click()
    Dim i As Integer

    If Check2.value = vbChecked Then

        With Me.VSFlexGrid2
 
            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
                .TextMatrix(i, .ColIndex("TransPayedValue")) = .TextMatrix(i, .ColIndex("RemainingValue"))
            Next i

        End With

    Else

        With Me.VSFlexGrid2

            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("payed")) = False
                .TextMatrix(i, .ColIndex("TransPayedValue")) = 0
            Next i

        End With

    End If
End Sub

Private Sub Check3_Click()
    Dim i As Integer

    If Check3.value = vbChecked Then

        With Me.VSFlexGrid3
 
            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
                .TextMatrix(i, .ColIndex("TransPayedValue")) = .TextMatrix(i, .ColIndex("RemainingValue"))
            Next i

        End With

    Else

        With Me.VSFlexGrid3

            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("payed")) = False
                .TextMatrix(i, .ColIndex("TransPayedValue")) = 0
            Next i

        End With

    End If
End Sub

Private Sub CMDSENDSMS_Click()
'0 manual
'1 save
'2 Print

SendMessage (0)
End Sub
Function SendMessage(currentOpt As Integer)
            Dim subject As String
            Dim Msg As String
            Dim msgstatus As Boolean
           Dim CompanyName As String
           Dim cOptions As ClsCompanyInfo
           Set cOptions = New ClsCompanyInfo
           Dim companyphone As String
           Dim opt As Integer
            Dim CurrentMessage As String
            Dim t As String
    CurrentMessage = ComposMessage(Me.Name, 0, "", "", opt)
  If opt = currentOpt And DCboCashType.ListIndex = 8 Then
  
      CompanyName = cOptions.ArabCompanyName '& CHR(13) & CurrentBranchName
     companyphone = cOptions.Company_Mobile
  '«·„«·ﬂ
 Msg = "  „ «” ·«„ﬂ„ „»·€   " & XPTxtVal.text & "  ··ÊÕœ… —ﬁ„   " & DcbUnitNo.text & CHR(13) & "    ··⁄ﬁ«— —ﬁ„ " & DcbIqara.text
t = sendMessageM("user", "password", Msg, "", GetCustomerNumber(DBCboClientName.BoundText))
 

DoEvents



MsgBox " „ «·«—”«·"
     
     
     End If
 
End Function
Function print_report3(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = " SELECT        dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value, dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
MySQL = MySQL & "                         dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Notes.akarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.Notes.OfficeValue,"
MySQL = MySQL & "                         dbo.Notes.Remark"
MySQL = MySQL & " FROM            dbo.Notes LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblAqar ON dbo.Notes.akarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
MySQL = MySQL & " Where (dbo.Notes.NoteID =" & val(XPTxtID.text) & ") and (dbo.Notes.NoteType = 5) "


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Expenses_order12Amolat.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Expenses_order12AmolatE.rpt"
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

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

     If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(4).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(4).AddCurrentValue GetRegVATNo(val(dcBranch.BoundText))
    End If
    
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
Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = " SELECT dbo.Notes.Note_Value2, dbo.Notes.PreBalaValue, dbo.Notes.PreBalaPayed, dbo.Notes.PreBalaRemain, dbo.Notes.PreBalaTransPyed, dbo.Notes.PreBalaNet,      dbo.Notes.NoteDate, dbo.Notes.NoteID, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value, dbo.Notes.NoteType, dbo.Notes.CusID, dbo.TblCustemers.CusName, "
MySQL = MySQL & "                      dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Notes.OfficeValue, dbo.Notes.RenterValue, dbo.Notes.ExpValue,"
MySQL = MySQL & "                      dbo.TblNotesOwnerPayment202.CusID AS Expr2, TblCustemers_1.CusName AS RenterName, TblCustemers_1.CusNamee AS RenterNameE,"
MySQL = MySQL & "                      TblCustemers_1.Fullcode AS RenterFullCode, dbo.TblNotesOwnerPayment202.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname,"
MySQL = MySQL & "                      dbo.TblNotesOwnerPayment202.UnitNo, dbo.TblAqarDetai.unitno AS UnitName, dbo.TblNotesOwnerPayment202.branch_no, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblNotesOwnerPayment202.TypTrans, dbo.TblNotesOwnerPayment202.Remarks,"
MySQL = MySQL & "                      dbo.TblNotesOwnerPayment202.ContNoteSerial1, dbo.TblNotesOwnerPayment202.[value], dbo.TblNotesOwnerPayment202.PayedValue,"
MySQL = MySQL & "                      dbo.TblNotesOwnerPayment202.RemainingValue, dbo.TblNotesOwnerPayment202.NetValue, dbo.TblNotesOwnerPayment202.TransPayedValue,"
MySQL = MySQL & "                      dbo.TblNotesOwnerPayment202.NoteID3, dbo.TblNotesOwnerPayment202.NoteDate AS DatePayed, dbo.Notes.TotalPayments,"
MySQL = MySQL & "                      NoteCashingType,TblCustemers.BankAccount BankName,BanksData.BankName as BankName2"
MySQL = MySQL & " FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblNotesOwnerPayment202 ON dbo.TblBranchesData.branch_id = dbo.TblNotesOwnerPayment202.branch_no LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAqarDetai ON dbo.TblNotesOwnerPayment202.UnitNo = dbo.TblAqarDetai.Id RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAqar ON dbo.TblNotesOwnerPayment202.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers TblCustemers_1 ON dbo.TblNotesOwnerPayment202.CusID = TblCustemers_1.CusID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.Notes ON dbo.TblNotesOwnerPayment202.NoteID = dbo.Notes.NoteID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID"
MySQL = MySQL & "                             LEFT OUTER JOIN dbo.BanksData"
MySQL = MySQL & "                                  ON  dbo.Notes.BankID = dbo.BanksData.BankID"
MySQL = MySQL & " Where (dbo.Notes.NoteID =" & val(XPTxtID.text) & ") and (dbo.Notes.NoteType = 5) "


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Expenses_order12.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Expenses_order12E.rpt"
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

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TotalPayments.text) + val(TxtTotalPayedOpBalance.text), "0.00"), 0, True, ".")
        xReport.ParameterFields(5).AddCurrentValue WriteNo(Format(val(TxtOfficeValueNet.text), "0.00"), 0, True, ".")
        xReport.ParameterFields(6).AddCurrentValue WriteNo(Format(val(TxtValuExpenses.text), "0.00"), 0, True, ".")
        xReport.ParameterFields(7).AddCurrentValue WriteNo(Format(val(TxtNetValue.text), "0.00"), 0, True, ".")
        xReport.ParameterFields(8).AddCurrentValue WriteNo(Format(val(TxtTotalPayedOpBalance.text), "0.00"), 0, True, ".")
        xReport.ParameterFields(9).AddCurrentValue WriteNo(Format(val(TotalPayments.text), "0.00"), 0, True, ".")
        

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

Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

MySQL = "SELECT     dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value2 as Note_Value, dbo.Notes.NoteDateH, "
MySQL = MySQL & "                      dbo.Notes.ContractNo, dbo.Notes.ContNo, dbo.Notes.commission, dbo.Notes.rent, dbo.Notes.Water, dbo.Notes.FilterID, dbo.Notes.FIlterTotal, dbo.Notes.Instrunce,"
MySQL = MySQL & "                      dbo.Notes.comX, dbo.Notes.ComY, dbo.Notes.CommissionOut, dbo.Notes.NoteOrBonID, dbo.Notes.comXold, dbo.Notes.ComYold, dbo.Notes.NoteOrBonValue,"
MySQL = MySQL & "                      dbo.Notes.NoteOrBonSereal, dbo.Notes.Telephone, dbo.Notes.CashingType, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
MySQL = MySQL & "                      dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.Notes.renterName, dbo.Notes.NoteCashingType, dbo.Notes.BankName, dbo.Notes.DueDate,"
MySQL = MySQL & "                      dbo.Notes.ChqueNum, dbo.Notes.Remark, dbo.Notes.Remark2, dbo.Notes.ToPriodDateH, dbo.Notes.FrmPriodDateH, dbo.Notes.ToPriodDate, dbo.Notes.FrmPriodDate,"
MySQL = MySQL & "                      dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAqarDetai.Id, dbo.TblAqarDetai.unitno,"
MySQL = MySQL & "                      dbo.TblAqarDetai.unittype, dbo.TblAqarDetai.Aqarid, dbo.TblAqar.aqarname, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.Notes.general_des_notes,"
MySQL = MySQL & "                      dbo.Notes.BankID, dbo.BanksData.BankName AS BankName1, dbo.BanksData.BankNamee, dbo.Notes.akarid, dbo.Notes.unittype AS unittype1,"
MySQL = MySQL & "                      dbo.Notes.EmpAccountCode, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name AS Emp_Name0, TblEmployee_1.Emp_Name1 AS Emp_Name10, TblEmployee_1.Emp_Code AS Emp_Code0,"
MySQL = MySQL & "                      TblEmployee_1.Emp_ID AS Emp_ID0, TblEmployee_1.Emp_Name2 AS Emp_Name20, TblEmployee_1.Emp_Name3 AS Emp_Name30,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name4 AS Emp_Name40, TblEmployee_1.Fullcode AS Fullcode0, TblEmployee_1.Emp_Namee4 AS Emp_Namee40,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee3 AS Emp_Namee30, TblEmployee_1.Emp_Namee2 AS Emp_Namee20, TblEmployee_1.Emp_Namee1 AS Emp_Namee10,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee AS Emp_Namee0, TblEmployee_2.Emp_ID AS Emp_ID1, TblEmployee_2.Emp_Code AS Emp_Code1,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Name AS Emp_Name_1, TblEmployee_2.Emp_Name1 AS Emp_Name11, TblEmployee_2.Emp_Name2 AS Emp_Name21,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Name3 AS Emp_Name31, TblEmployee_2.Emp_Name4 AS Emp_Name41, TblEmployee_2.Fullcode AS Fullcode1,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Namee4 AS Emp_Namee41, TblEmployee_2.Emp_Namee3 AS Emp_Namee31, TblEmployee_2.Emp_Namee2 AS Emp_Namee21,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Namee1 AS Emp_Namee11, TblEmployee_2.Emp_Namee AS Emp_Namee_1, TblEmployee_3.Emp_ID AS Emp_ID2,"
MySQL = MySQL & "                      TblEmployee_3.Emp_Code AS Emp_Code2, TblEmployee_3.Emp_Name AS Emp_Name_2, TblEmployee_3.Emp_Name1 AS Emp_Name12,"
MySQL = MySQL & "                      TblEmployee_3.Emp_Name2 AS Emp_Name22, TblEmployee_3.Emp_Name3 AS Emp_Name32, TblEmployee_3.Emp_Name4 AS Emp_Name42,"
MySQL = MySQL & "                      TblEmployee_3.Fullcode AS Fullcode2, TblEmployee_3.Emp_Namee4 AS Emp_Namee42, TblEmployee_3.Emp_Namee3 AS Emp_Namee32,"
MySQL = MySQL & "                      TblEmployee_3.Emp_Namee2 AS Emp_Namee22, TblEmployee_3.Emp_Namee1 AS Emp_Namee12, TblEmployee_3.Emp_Namee AS Emp_Namee_2,"
MySQL = MySQL & "                      dbo.Notes.BTCashAccountcode , dbo.ACCOUNTS.Account_Code, dbo.ACCOUNTS.account_name, dbo.ACCOUNTS.Account_NameEng"
MySQL = MySQL & " FROM         dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_3 RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.ACCOUNTS RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.Notes ON dbo.ACCOUNTS.Account_Code = dbo.Notes.BTCashAccountcode ON TblEmployee_3.Account_Code3 = dbo.Notes.EmpAccountCode LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.Notes.EmpAccountCode = TblEmployee_2.Account_code1 ON"
MySQL = MySQL & "                      TblEmployee_1.Account_code = dbo.Notes.EmpAccountCode LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.Notes.EmpAccountCode = dbo.TblEmployee.Account_Code2 LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAkarUnit ON dbo.Notes.unittype = dbo.TblAkarUnit.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAqar ON dbo.Notes.akarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAqarDetai ON dbo.Notes.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID"

MySQL = MySQL & " Where (dbo.Notes.NoteID =" & val(XPTxtID.text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Expenses_order11.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Expenses_order11.rpt"
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

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""

    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(XPTxtVal.text), "0.00"), 0, True, ".")

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

Function CheckStatusofUnit(ID As Double) As Boolean

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Dim sql As String
        Dim Status As Boolean
        Dim i As Integer
        Dim rs As New ADODB.Recordset
 Status = True
        sql = " SELECT    StatusEarnest "
        sql = sql & " from dbo.TblAqrEarnest"
        sql = sql & " Where (UnitNo = " & ID & ") "

        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.RecordCount > 0 Then
        
        For i = 1 To rs.RecordCount
        If rs("StatusEarnest").value = 0 Then
        Status = False
     End If
        
  Next i
    End If
    CheckStatusofUnit = Status
End If
End Function

Sub GetNotesSalesInformation(Optional NodID As Double)
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL As String
 Dim i, emp As Integer
 Dim Rate As Double

       Set RsDetails1 = New ADODB.Recordset
         StrSQL = " SELECT     dbo.Notes.NoteID, dbo.TblNotesSales.idd, dbo.TblNotesSales.Type, dbo.TblNotesSales.valu, dbo.TblNotesSales.EmpID, dbo.TblNotesSales.rate,"
         StrSQL = StrSQL & "             dbo.TblNotesSales.id , dbo.Notes.CashingType"
         StrSQL = StrSQL & " FROM         dbo.Notes LEFT OUTER JOIN"
         StrSQL = StrSQL & "                    dbo.TblNotesSales ON dbo.Notes.NoteID = dbo.TblNotesSales.NoteID"
         StrSQL = StrSQL & " Where (CashingType = 9) And (dbo.Notes.NoteID =" & NodID & ")"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails1.RecordCount > 0 Then
   RsDetails1.MoveFirst
   For i = 1 To RsDetails1.RecordCount
   Rate = IIf(IsNull(RsDetails1("rate").value), 0, Trim(RsDetails1("rate").value))
   emp = IIf(IsNull(RsDetails1("idd").value), 0, Trim(RsDetails1("idd").value))
   AqrCommisiion val(DCboCashType.ListIndex), emp, Rate
   
   RsDetails1.MoveNext
  Next i
    End If
   End Sub
   
Sub GetNotesInformation(Optional NodID As Double)
 Dim RsDetails1 As ADODB.Recordset
 Dim StrSQL As String

       Set RsDetails1 = New ADODB.Recordset
         StrSQL = " SELECT     dbo.Notes.*"
StrSQL = StrSQL & " From dbo.Notes"
StrSQL = StrSQL & " Where (CashingType = 9) And (NoteID =" & NodID & ")"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If RsDetails1.RecordCount > 0 Then
      TxtRent.text = IIf(IsNull(RsDetails1("rent").value), "", Trim(RsDetails1("rent").value))
      Txtcommission.text = IIf(IsNull(RsDetails1("commission").value), "", Trim(RsDetails1("commission").value))
        Me.TxtCommissionOut.text = IIf(IsNull(RsDetails1("CommissionOut").value), "", Trim(RsDetails1("CommissionOut").value))
       TxtWater.text = IIf(IsNull(RsDetails1("Water").value), "", Trim(RsDetails1("Water").value))
        txtinstrunce.text = IIf(IsNull(RsDetails1("Instrunce").value), "", Trim(RsDetails1("Instrunce").value))
         txtComisinold.text = IIf(IsNull(RsDetails1("comX").value), "", Trim(RsDetails1("comX").value))
          txtinstrancold.text = IIf(IsNull(RsDetails1("ComY").value), "", Trim(RsDetails1("ComY").value))
    End If
   End Sub
Sub AqrCommisiion(Optional index As Integer, Optional emp As Integer, Optional Rate As Double)
 Dim RsDetails1 As ADODB.Recordset
Dim comx As Integer
Dim comy As Integer
 Dim StrSQL, sql As String

       Set RsDetails1 = New ADODB.Recordset
         StrSQL = "SELECT     *  from  TblAqarCommissions Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
Select Case index

Case 7

If val(Me.txtDiff.text) = 0 Then
comx = val(Me.txtComisinold.text)
comy = val(Me.txtinstrancold.text)
  sql = "update Notes set   StatusEarnest =2   where  NoteID =" & val(Me.TxtNotID.text) & " "
        Cn.Execute sql
        
       If CheckStatusofUnit(val(Me.DcbUnitNo.BoundText)) = True Then
       
        sql = "update TblAqarDetai set   Status =0   where  Id =" & val(Me.DcbUnitNo.BoundText) & " "
        Cn.Execute sql
       End If
       
Else
 sql = "update Notes set   StatusEarnest =3   where  NoteID =" & val(Me.TxtNotID.text) & " "
        Cn.Execute sql
comx = val(Me.txtComisinold.text) - val(txtComisin.text)
comy = val(Me.txtinstrancold.text) - val(txtinstranc.text)
End If

If comx <> 0 Then
           RsDetails1.AddNew
           RsDetails1("PymentNo").value = Null
           RsDetails1("ContNo").value = Null
           RsDetails1("FilterNo").value = Null
           RsDetails1("NoteID").value = val(XPTxtID.text)
           RsDetails1("IqarID").value = val(Me.DcbUnitNo.BoundText)
           RsDetails1("TypeOper").value = 9
           RsDetails1("TypeAmount").value = 1
           RsDetails1("EmpID").value = emp
           RsDetails1("Amount").value = (Rate / 100) * comx * -1
                RsDetails1.update
   End If
       
         '''\\\\\
     If comy <> 0 Then
            RsDetails1.AddNew
           RsDetails1("PymentNo").value = Null
           RsDetails1("ContNo").value = Null
           RsDetails1("FilterNo").value = Null
           RsDetails1("NoteID").value = val(XPTxtID.text)
           RsDetails1("IqarID").value = val(Me.DcbUnitNo.BoundText)
           RsDetails1("TypeOper").value = 9
           RsDetails1("TypeAmount").value = 2
           RsDetails1("EmpID").value = emp
           RsDetails1("Amount").value = val(Rate / 100) * comy * -1
           RsDetails1.update
  
End If
End Select

End Sub

Private Sub ALLButton1_Click()

    If IsNumeric(Me.DBCboClientName.BoundText) Then
        'INSTALLMENT_DATA2.show
        'INSTALLMENT_DATA2.Adodc1.CommandType = adCmdText
        'INSTALLMENT_DATA2.Adodc1.RecordSource = "select *  FROM INSTALLMENT_DETAILS where payed=1 and  cust_id =" & Me.DBCboClientName.BoundText
        'INSTALLMENT_DATA2.Adodc1.Refresh
 '
 '       INSTALLMENT_DATA2.id.text = Me.DBCboClientName.BoundText
 '       INSTALLMENT_DATA2.lblcustid = Me.DBCboClientName.BoundText
 '       INSTALLMENT_DATA2.TxtName.text = Me.DBCboClientName.text
    End If

End Sub

Private Sub ALLButton2_Click()

    If IsNumeric(Me.DBCboClientName.BoundText) Then
        'sanad_dean.show
        'sanad_dean.LblID = DBCboClientName.BoundText
        'sanad_dean.LblName = DBCboClientName.text
        'sanad_dean.lblaccountcode.Caption = txtaccount.text
        'sanad_dean.Adodc1.CommandType = adCmdText
        'sanad_dean.Adodc1.RecordSource = "select*  FROM sanad_dean where cust_id=" & DBCboClientName.BoundText
        'sanad_dean.Adodc1.Refresh
        'sanad_dean.ALLButton1.Visible = False
        'sanad_dean.ALLButton1.Visible = False
'
'        sanad_dean.Adodc2.CommandType = adCmdText
'        sanad_dean.Adodc2.RecordSource = "select *  FROM member_child where cust_id=" & DBCboClientName.BoundText
'        sanad_dean.Adodc2.Refresh
    End If

End Sub

Private Sub ALLButton3_Click()
    lblsqlstring.Caption = ""
  '  FrmPaymentTime2.show
  '  FrmPaymentTime2.lblcusid = DBCboClientName.BoundText
  '  FrmPaymentTime2.LblValue = val(XPTxtVal.Text)

If val(DBCboClientName.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— «·„«·ﬂ «Ê·«"
Else
MsgBox "Please Select Owner"
End If
DBCboClientName.SetFocus
Exit Sub
Else
Frame12(0).Visible = True
If Me.TxtModFlg.text <> "R" Then
XPTxtVal.text = 0

If Me.TxtModFlg.text = "N" Then
RetriveOwnerPayment val(DBCboClientName.BoundText)
End If

If Me.TxtModFlg.text = "E" And (FlgBillBuy = True Or VSFlexGrid1.rows = 1) Then
RetriveOwnerPayment val(DBCboClientName.BoundText)
End If
End If
End If

End Sub
Sub RetriveOwnerPayment(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
Set Rs8 = New ADODB.Recordset
sql = " SELECT     dbo.TblAqar.Aqarid, dbo.TblAqar.ownerid, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblAqar.aqarname,"
sql = sql & "                      dbo.TblAqar.aqarNo, dbo.TblAqar.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAqrOwin.RecDateH,"
sql = sql & "                      dbo.TblAqrOwin.RecDate, dbo.TblAqrOwin.[value], dbo.TblAqrOwin.DMY, dbo.TblAqrOwin.Cont, dbo.TblAqrOwin.AllowDateH, dbo.TblAqrOwin.AllowDate,"
sql = sql & "                      dbo.TblAqrOwin.PaymentNo , dbo.TblAqrOwin.ID"
sql = sql & " FROM         dbo.TblAqar LEFT OUTER JOIN"
sql = sql & "                      dbo.TblAqrOwin ON dbo.TblAqar.Aqarid = dbo.TblAqrOwin.AqrID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.TblAqar.ownerid = dbo.TblCustemers.CusID"
sql = sql & " where dbo.TblAqar.ownerid=" & val(CuID) & " and ( dbo.TblAqrOwin.ID<>0 or not( dbo.TblAqrOwin.ID is null))and  (dbo.TblAqrOwin.TotalPayed = 0 or dbo.TblAqrOwin.TotalPayed is null)"
If val(DcbIqara2.BoundText) <> 0 And Trim(DcbIqara2.text) <> "" Then
    sql = sql & " and TblAqrOwin.AqrID = " & val(DcbIqara2.BoundText)
End If
sql = sql & "  ORDER BY dbo.TblAqrOwin.RecDate"
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
.TextMatrix(i, .ColIndex("Aqarid")) = IIf(IsNull(Rs8("Aqarid").value), 0, Rs8("Aqarid").value)
.TextMatrix(i, .ColIndex("aqarname")) = IIf(IsNull(Rs8("aqarname").value), "", Rs8("aqarname").value)
.TextMatrix(i, .ColIndex("aqarNo")) = IIf(IsNull(Rs8("aqarNo").value), "", Rs8("aqarNo").value)
.TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(Rs8("BranchId").value), 0, Rs8("BranchId").value)
.TextMatrix(i, .ColIndex("RecDateH")) = IIf(IsNull(Rs8("RecDateH").value), "", Rs8("RecDateH").value)
.TextMatrix(i, .ColIndex("RecDate")) = IIf(IsNull(Rs8("RecDate").value), "", Rs8("RecDate").value)
.TextMatrix(i, .ColIndex("AllowDateH")) = IIf(IsNull(Rs8("AllowDateH").value), "", Rs8("AllowDateH").value)
.TextMatrix(i, .ColIndex("AllowDate")) = IIf(IsNull(Rs8("AllowDate").value), "", Rs8("AllowDate").value)
.TextMatrix(i, .ColIndex("PaymentNo")) = IIf(IsNull(Rs8("PaymentNo").value), "", Rs8("PaymentNo").value)
.TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(Rs8("ID").value), 0, Rs8("ID").value)
.TextMatrix(i, .ColIndex("value")) = IIf(IsNull(Rs8("value").value), 0, Rs8("value").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
Else
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
End If
If val(.TextMatrix(i, .ColIndex("PaymentID"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteOwnerPayed(val(.TextMatrix(i, .ColIndex("PaymentID"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
Rs8.MoveNext
Next i
End With
End If

End Sub

Function GetInsuranceValue(Optional NoteID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     isnull(insurancepayed,0) as value"
sql = sql & " From dbo.ContracttBillInstallmentsDone"
sql = sql & " WHERE     (NoteID = " & NoteID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetInsuranceValue = IIf(IsNull(rs2("value").value), 0, rs2("value").value)
Else
GetInsuranceValue = 0
End If
End Function


Function GetCommitionValue(Optional NoteID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     (isnull(CommissionsPayed,0)+ isnull(0,0)+ isnull(TelandNetPayed,0)) as value"
sql = sql & " From dbo.ContracttBillInstallmentsDone"
sql = sql & " WHERE     (NoteID = " & NoteID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetCommitionValue = IIf(IsNull(rs2("value").value), 0, rs2("value").value)
Else
GetCommitionValue = 0
End If
End Function
Sub RetriveOwnerPayment2031(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
Dim CurrRow As Integer
Set Rs8 = New ADODB.Recordset
sql = " SELECT        dbo.Notes.NoteID, dbo.Notes.Insurance, dbo.Notes.CusID, dbo.Notes.NoteType, dbo.Notes.branch_no, dbo.Notes.NoteDate, dbo.Notes.Remark2, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCustemers.CusName, "
sql = sql & "                         dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Notes.akarid, dbo.TblAqar.aqarname, dbo.Notes.FilterID2, dbo.Notes.RemainWater, dbo.Notes.BillPrice, dbo.Notes.RemainCommissions,"
sql = sql & "                         dbo.Notes.OldRent, dbo.Notes.RemainService, dbo.Notes.txtOldInsurance, dbo.Notes.RemainRent, dbo.Notes.Instrunce, dbo.TblAqar.ownerid, dbo.Notes.UnitNo, dbo.TblAqarDetai.unitno AS UntName,"
sql = sql & "                         dbo.TblNotesTypes.NotesTypeName , dbo.TblNotesTypes.NotesTypeNameE"
sql = sql & " FROM            dbo.Notes INNER JOIN"
sql = sql & "                         dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType LEFT OUTER JOIN"
sql = sql & "                         dbo.TblAqarDetai ON dbo.Notes.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
sql = sql & "                         dbo.TblAqar ON dbo.Notes.akarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
sql = sql & "                         dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                         dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
If SystemOptions.EndRentifPayed = True Then
sql = sql & " WHERE        (dbo.Notes.NoteType = - 1)  "
sql = sql & " and dbo.TblAqar.Aqarid =" & val(Me.DcbIqara2.BoundText) & " and  dbo.TblAqar.ownerid =" & val(CuID) & " and   (dbo.Notes.TotalPayed <> 0 AND NOT( dbo.Notes.TotalPayed is null))"


Else
sql = sql & " WHERE        (dbo.Notes.NoteType = - 1) AND (dbo.Notes.totalPayed =  0 OR"
sql = sql & "                         dbo.Notes.totalPayed IS NULL)"
sql = sql & " and dbo.TblAqar.Aqarid =" & val(Me.DcbIqara2.BoundText) & " and  dbo.TblAqar.ownerid =" & val(CuID) & " and   (dbo.Notes.TotalPayed = 0 or dbo.Notes.TotalPayed is null)"
End If

Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
VSFlexGrid3.Enabled = True
        VSFlexGrid3.Enabled = True
With VSFlexGrid3
CurrRow = .rows
    .rows = .rows + Rs8.RecordCount
'.Rows = .FixedRows + Rs8.RecordCount
Rs8.MoveFirst
For i = CurrRow To .rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("NoteType")) = IIf(IsNull(Rs8("NoteType").value), 0, Rs8("NoteType").value)
.TextMatrix(i, .ColIndex("NotesTypeName")) = " ’›Ì… "
.TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(Rs8("CusID").value), 0, Rs8("CusID").value)
.TextMatrix(i, .ColIndex("UnitNo")) = IIf(IsNull(Rs8("UnitNo").value), 0, Rs8("UnitNo").value)
.TextMatrix(i, .ColIndex("Aqarid")) = IIf(IsNull(Rs8("akarid").value), 0, Rs8("akarid").value)
.TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("branch_no").value), 0, Rs8("branch_no").value)
.TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs8("Remark2").value), "", Rs8("Remark2").value)
.TextMatrix(i, .ColIndex("aqarname")) = IIf(IsNull(Rs8("aqarname").value), "", Rs8("aqarname").value)
.TextMatrix(i, .ColIndex("UntName")) = IIf(IsNull(Rs8("UntName").value), "", Rs8("UntName").value)
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("NoteDate").value), "", Rs8("NoteDate").value)
.TextMatrix(i, .ColIndex("NoteID2")) = IIf(IsNull(Rs8("NoteID").value), "", Rs8("NoteID").value)
.TextMatrix(i, .ColIndex("ContNoteSerial1")) = IIf(IsNull(Rs8("FilterID2").value), "", Rs8("FilterID2").value)
.TextMatrix(i, .ColIndex("value")) = IIf(IsNull(Rs8("RemainRent").value), 0, Rs8("RemainWater").value) + IIf(IsNull(Rs8("RemainWater").value), 0, Rs8("RemainRent").value) + IIf(IsNull(Rs8("BillPrice").value), 0, Rs8("BillPrice").value) + IIf(IsNull(Rs8("insurance").value), 0, Rs8("insurance").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
.TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs8("CusName").value), "", Rs8("CusName").value)
'.TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(Rs8("NotesTypeName").value), "", Rs8("NotesTypeName").value)
Else
'.TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(Rs8("NotesTypeNamee").value), "", Rs8("NotesTypeNamee").value)
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
.TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
End If
If val(.TextMatrix(i, .ColIndex("NoteID2"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteOwnerPayed202(val(.TextMatrix(i, .ColIndex("NoteID2"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
Rs8.MoveNext
Next i
End With
End If

End Sub
Sub RetriveOwnerPayment2021(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
Dim CurrRow As Integer
Set Rs8 = New ADODB.Recordset
sql = " SELECT        dbo.Notes.NoteID ,dbo.Notes.NoteSerial1,dbo.Notes.Note_Value2, dbo.Notes.Insurance, dbo.Notes.CusID, dbo.Notes.NoteType, dbo.Notes.branch_no, dbo.Notes.NoteDate, dbo.Notes.Remark2, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCustemers.CusName, "
sql = sql & "                         dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Notes.akarid, dbo.TblAqar.aqarname, dbo.Notes.FilterID2, dbo.Notes.RemainWater, dbo.Notes.BillPrice, dbo.Notes.RemainCommissions,"
sql = sql & "                         dbo.Notes.OldRent, dbo.Notes.RemainService, dbo.Notes.txtOldInsurance, dbo.Notes.RemainRent, dbo.Notes.Instrunce, dbo.TblAqar.ownerid, dbo.Notes.UnitNo, dbo.TblAqarDetai.unitno AS UntName,"
sql = sql & "                         dbo.TblNotesTypes.NotesTypeName , dbo.TblNotesTypes.NotesTypeNameE"
sql = sql & " FROM            dbo.Notes INNER JOIN"
sql = sql & "                         dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType LEFT OUTER JOIN"
sql = sql & "                         dbo.TblAqarDetai ON dbo.Notes.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
sql = sql & "                         dbo.TblAqar ON dbo.Notes.akarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
sql = sql & "                         dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                         dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
sql = sql & " WHERE        (CashingType = 9 and NoteType=4) AND (dbo.Notes.totalPayed = 0 OR"
sql = sql & "                         dbo.Notes.totalPayed IS NULL)"
sql = sql & " and dbo.TblAqar.Aqarid =" & val(Me.DcbIqara2.BoundText) & " and  dbo.TblAqar.ownerid =" & val(CuID) & " and   (dbo.Notes.TotalPayed = 0 or dbo.Notes.TotalPayed is null)"

Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
VSFlexGrid2.Enabled = True
        VSFlexGrid2.Enabled = True
With VSFlexGrid2
CurrRow = .rows
    .rows = .rows + Rs8.RecordCount
'.Rows = .FixedRows + Rs8.RecordCount
Rs8.MoveFirst
For i = CurrRow To .rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("NoteType")) = 9
.TextMatrix(i, .ColIndex("NotesTypeName")) = "⁄—»Ê‰"
.TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(Rs8("CusID").value), 0, Rs8("CusID").value)
.TextMatrix(i, .ColIndex("UnitNo")) = IIf(IsNull(Rs8("UnitNo").value), 0, Rs8("UnitNo").value)
.TextMatrix(i, .ColIndex("Aqarid")) = IIf(IsNull(Rs8("akarid").value), 0, Rs8("akarid").value)
.TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("branch_no").value), 0, Rs8("branch_no").value)
.TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs8("Remark2").value), "", Rs8("Remark2").value)
.TextMatrix(i, .ColIndex("aqarname")) = IIf(IsNull(Rs8("aqarname").value), "", Rs8("aqarname").value)
.TextMatrix(i, .ColIndex("UntName")) = IIf(IsNull(Rs8("UntName").value), "", Rs8("UntName").value)
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("NoteDate").value), "", Rs8("NoteDate").value)
.TextMatrix(i, .ColIndex("NoteID2")) = IIf(IsNull(Rs8("NoteID").value), "", Rs8("NoteID").value)
.TextMatrix(i, .ColIndex("ContNoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
.TextMatrix(i, .ColIndex("value")) = IIf(IsNull(Rs8("Note_Value2").value), 0, Rs8("Note_Value2").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
.TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs8("CusName").value), "", Rs8("CusName").value)

Else

.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
.TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
End If
If val(.TextMatrix(i, .ColIndex("NoteID2"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteOwnerPayed202(val(.TextMatrix(i, .ColIndex("NoteID2"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
Rs8.MoveNext
Next i
End With
End If

End Sub
Sub RetriveOwnerPayment202(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
With VSFlexGrid2
.Clear flexClearScrollable, flexClearEverything
.rows = 1
End With
Set Rs8 = New ADODB.Recordset
sql = " SELECT    dbo.Notes.noteserial1,     dbo.Notes.NoteID, dbo.Notes.ContNo, dbo.Notes.NoteType, dbo.Notes.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblContract.CusID, dbo.TblCustemers.CusName, "
sql = sql & "                         dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Notes.NoteDate, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value2, dbo.TblContract.Iqar, dbo.TblAqar.aqarname, dbo.TblAqar.aqarNo,"
sql = sql & "                         dbo.TblContract.NoteSerial1 AS ContNoteSerial1, dbo.TblContract.UnitNo, dbo.TblAqarDetai.unitno AS UntName, dbo.Notes.Remark2, dbo.TblContract.ownerid, dbo.TblNotesTypes.NotesTypeName,"
sql = sql & "                         dbo.TblNotesTypes.NotesTypeNameE"
sql = sql & " FROM            dbo.Notes INNER JOIN"
sql = sql & "                         dbo.TblContract ON dbo.Notes.ContNo = dbo.TblContract.ContNo INNER JOIN"
sql = sql & "                         dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType LEFT OUTER JOIN"
sql = sql & "                         dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
sql = sql & "                         dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
sql = sql & "                         dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                         dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
sql = sql & " Where (dbo.Notes.NoteType = 4)"
sql = sql & " and dbo.TblAqar.Aqarid =" & val(Me.DcbIqara2.BoundText) & " and  dbo.TblContract.ownerid =" & val(CuID) & " and   (dbo.Notes.TotalPayed = 0 or dbo.Notes.TotalPayed is null)"

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
 
 .TextMatrix(i, .ColIndex("noteserial1")) = IIf(IsNull(Rs8("noteserial1").value), "", Rs8("noteserial1").value)
 
.TextMatrix(i, .ColIndex("NoteType")) = IIf(IsNull(Rs8("NoteType").value), 0, Rs8("NoteType").value)
.TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(Rs8("CusID").value), 0, Rs8("CusID").value)
.TextMatrix(i, .ColIndex("UnitNo")) = IIf(IsNull(Rs8("UnitNo").value), 0, Rs8("UnitNo").value)
.TextMatrix(i, .ColIndex("Aqarid")) = IIf(IsNull(Rs8("Iqar").value), 0, Rs8("Iqar").value)
.TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("branch_no").value), 0, Rs8("branch_no").value)
.TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs8("Remark2").value), "", Rs8("Remark2").value)
.TextMatrix(i, .ColIndex("aqarname")) = IIf(IsNull(Rs8("aqarname").value), "", Rs8("aqarname").value)
.TextMatrix(i, .ColIndex("UntName")) = IIf(IsNull(Rs8("UntName").value), "", Rs8("UntName").value)
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("NoteDate").value), "", Rs8("NoteDate").value)
.TextMatrix(i, .ColIndex("NoteID2")) = IIf(IsNull(Rs8("NoteID").value), "", Rs8("NoteID").value)
.TextMatrix(i, .ColIndex("NotesTypeName")) = "„ﬁ»Ê÷« "
.TextMatrix(i, .ColIndex("ContNoteSerial1")) = IIf(IsNull(Rs8("ContNoteSerial1").value), "", Rs8("ContNoteSerial1").value)
.TextMatrix(i, .ColIndex("value")) = IIf(IsNull(Rs8("Note_Value2").value), 0, Rs8("Note_Value2").value) - GetCommitionValue(val(.TextMatrix(i, .ColIndex("NoteID2"))))
.TextMatrix(i, .ColIndex("InsuranceValue")) = GetInsuranceValue(val(.TextMatrix(i, .ColIndex("NoteID2"))))


If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
.TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs8("CusName").value), "", Rs8("CusName").value)
'.TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(Rs8("NotesTypeName").value), "", Rs8("NotesTypeName").value)
Else
'.TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(Rs8("NotesTypeNamee").value), "", Rs8("NotesTypeNamee").value)
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
.TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs8("branch_namee").value), "", Rs8("branch_namee").value)
End If
If val(.TextMatrix(i, .ColIndex("NoteID2"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteOwnerPayed202(val(.TextMatrix(i, .ColIndex("NoteID2"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
Rs8.MoveNext
Next i
End With
End If
RetriveOwnerPayment2021 CuID
End Sub
Sub RetriveOwnerPayment203(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
With VSFlexGrid3
.Clear flexClearScrollable, flexClearEverything
.rows = 1
End With
Set Rs8 = New ADODB.Recordset
sql = " SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteSerial1, dbo.notes_all.branch_no, dbo.TblBranchesData.branch_name, "
sql = sql & "                      dbo.TblBranchesData.branch_namee, dbo.notes_all.too, dbo.notes_all.NoteType, dbo.TblExpensesDet.iqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname,"
sql = sql & "                      dbo.TblExpensesDet.des,dbo.TblExpensesDet.ID ,dbo.TblExpUnitNo.ID as IDDet, dbo.TblExpUnitNo.Valu, dbo.TblExpUnitNo.UnitID,dbo.TblExpensesDet.value, dbo.TblAqarDetai.unitno, dbo.TblAqar.ownerid"
sql = sql & " FROM         dbo.TblExpUnitNo LEFT OUTER JOIN"
sql = sql & "                      dbo.TblAqarDetai ON dbo.TblExpUnitNo.UnitID = dbo.TblAqarDetai.Id RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblExpensesDet ON dbo.TblExpUnitNo.ExpDetails = dbo.TblExpensesDet.ID RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblAqar ON dbo.TblExpensesDet.iqarid = dbo.TblAqar.Aqarid RIGHT OUTER JOIN"
sql = sql & "                      dbo.notes_all ON dbo.TblExpensesDet.ExpID = dbo.notes_all.NoteID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
sql = sql & " Where (dbo.notes_all.NoteType = 3) "
sql = sql & " and dbo.TblAqar.Aqarid=" & val(Me.DcbIqara2.BoundText) & " and  dbo.TblAqar.ownerid =" & val(CuID) & " and   (dbo.TblExpensesDet.TotalPayed = 0 or dbo.TblExpensesDet.TotalPayed is null)"

Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
VSFlexGrid3.Enabled = True
        VSFlexGrid3.Enabled = True
With VSFlexGrid3
.Clear flexClearScrollable, flexClearEverything
.rows = 1
    .rows = .rows + Rs8.RecordCount
.rows = .FixedRows + Rs8.RecordCount
Rs8.MoveFirst
For i = .FixedRows To Rs8.RecordCount
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("NoteDate").value), "", Rs8("NoteDate").value)
.TextMatrix(i, .ColIndex("NoteID2")) = IIf(IsNull(Rs8("ID").value), "", Rs8("ID").value)
.TextMatrix(i, .ColIndex("NoteID3")) = IIf(IsNull(Rs8("IDDet").value), "", Rs8("IDDet").value)
.TextMatrix(i, .ColIndex("ContNoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
.TextMatrix(i, .ColIndex("UnitNo")) = IIf(IsNull(Rs8("UnitID").value), 0, Rs8("UnitID").value)
.TextMatrix(i, .ColIndex("Aqarid")) = IIf(IsNull(Rs8("iqarid").value), 0, Rs8("iqarid").value)
.TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("branch_no").value), 0, Rs8("branch_no").value)
.TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs8("des").value), "", Rs8("des").value)
.TextMatrix(i, .ColIndex("aqarname")) = IIf(IsNull(Rs8("aqarname").value), "", Rs8("aqarname").value)
.TextMatrix(i, .ColIndex("UntName")) = IIf(IsNull(Rs8("unitno").value), "", Rs8("unitno").value)
.TextMatrix(i, .ColIndex("branch_name")) = "„’—Ê›« "
.TextMatrix(i, .ColIndex("value")) = IIf(IsNull(Rs8("Valu").value), IIf(IsNull(Rs8("value").value), 0, Rs8("value").value), Rs8("Valu").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
Else
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
End If
If val(.TextMatrix(i, .ColIndex("NoteID2"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteOwnerPayed203(val(.TextMatrix(i, .ColIndex("NoteID2")))) + GeteOwnerPayed204(val(.TextMatrix(i, .ColIndex("NoteID3"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
Rs8.MoveNext
Next i
End With
End If
RetriveOwnerPayment2031 val(DBCboClientName.BoundText)
End Sub
Function SaveOwnerPayment()
    Dim StrSQL As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Diff = 0
   Dim RsDetails As ADODB.Recordset
   If Me.TxtModFlg.text = "E" Then
    StrSQL = " Delete From TblNotesOwnerPayment Where NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblOwnerPayment Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     * from dbo.TblNotesOwnerPayment Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
    'TxtValueTemp.Text = val(XPTxtVal.Text)
    For i = .FixedRows To .rows - 1
    TxtValueTemp = 0
        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.text)
            RsDetails("PaymentID").value = val(.TextMatrix(i, .ColIndex("PaymentID")))
            RsDetails("BranchId").value = val(.TextMatrix(i, .ColIndex("BranchId")))
            RsDetails("Aqarid").value = val(.TextMatrix(i, .ColIndex("Aqarid")))
            RsDetails("value").value = val(.TextMatrix(i, .ColIndex("value")))
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
           ' .TextMatrix(i, .ColIndex("TransPayedValue")) = DIFF
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
            RsDetails("RecDateH").value = (.TextMatrix(i, .ColIndex("RecDateH")))
            RsDetails("RecDate").value = IIf(.TextMatrix(i, .ColIndex("RecDate")) = "", Null, .TextMatrix(i, .ColIndex("RecDate")))
            RsDetails("AllowDateH").value = (.TextMatrix(i, .ColIndex("AllowDateH")))
            RsDetails("AllowDate").value = IIf(.TextMatrix(i, .ColIndex("AllowDate")) = "", Null, .TextMatrix(i, .ColIndex("AllowDate")))
            RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
            RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            RsDetails.update
                
            If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
            StrSQL = "Update TblAqrOwin Set  TotalPayed=1 Where ID=" & val(.TextMatrix(i, .ColIndex("PaymentID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
             Else
                 StrSQL = "Update TblAqrOwin Set  TotalPayed=0 Where ID=" & val(.TextMatrix(i, .ColIndex("PaymentID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
      End If
    Next i
End With
    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     * from dbo.TblOwnerPayment Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
    For i = .FixedRows To .rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.text)
            RsDetails("PaymentID").value = val(.TextMatrix(i, .ColIndex("PaymentID")))
            RsDetails("value").value = val(.TextMatrix(i, .ColIndex("value")))
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails.update
        End If
    Next i
End With

End Function
Function SaveOwnerPayment202()
    Dim StrSQL As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Diff = 0
   Dim RsDetails As ADODB.Recordset
   If Me.TxtModFlg.text = "E" Then
    StrSQL = " Delete From TblNotesOwnerPayment202 Where NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblOwnerPayment202 Where   NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     * from dbo.TblNotesOwnerPayment202 Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid2
    For i = .FixedRows To .rows - 1
    TxtValueTemp = 0
        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.text)
            RsDetails("NoteID2").value = val(.TextMatrix(i, .ColIndex("NoteID2")))
            RsDetails("NoteType").value = val(.TextMatrix(i, .ColIndex("NoteType")))
            RsDetails("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
            RsDetails("UnitNo").value = val(.TextMatrix(i, .ColIndex("UnitNo")))
            RsDetails("Aqarid").value = val(.TextMatrix(i, .ColIndex("Aqarid")))
            RsDetails("branch_no").value = val(.TextMatrix(i, .ColIndex("branch_no")))
            RsDetails("ContNoteSerial1").value = (.TextMatrix(i, .ColIndex("ContNoteSerial1")))
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
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
            RsDetails("NoteDate").value = IIf(.TextMatrix(i, .ColIndex("NoteDate")) = "", Null, .TextMatrix(i, .ColIndex("NoteDate")))
            RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
            RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            RsDetails("TypTrans").value = 0
            RsDetails.update
                
            If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
            StrSQL = "Update Notes Set  TotalPayed=1 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID2"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
             Else
                 StrSQL = "Update Notes Set  TotalPayed=0 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID2"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
      End If
    Next i
End With
    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     * from dbo.TblOwnerPayment202 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid2
    For i = .FixedRows To .rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.text)
            RsDetails("NoteID2").value = val(.TextMatrix(i, .ColIndex("NoteID2")))
            RsDetails("NoteType").value = val(.TextMatrix(i, .ColIndex("NoteType")))
            RsDetails("value").value = val(.TextMatrix(i, .ColIndex("value")))
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("TypTrans").value = 0
            RsDetails.update
        End If
    Next i
End With

''////////////////////
    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     * from dbo.TblNotesOwnerPayment202 Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid3
    For i = .FixedRows To .rows - 1
    TxtValueTemp = 0
        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.text)
             RsDetails("NoteType").value = val(.TextMatrix(i, .ColIndex("NoteType")))
            RsDetails("NoteID2").value = val(.TextMatrix(i, .ColIndex("NoteID2")))
            RsDetails("NoteID3").value = val(.TextMatrix(i, .ColIndex("NoteID3")))
            RsDetails("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
            RsDetails("UnitNo").value = val(.TextMatrix(i, .ColIndex("UnitNo")))
            RsDetails("Aqarid").value = val(.TextMatrix(i, .ColIndex("Aqarid")))
            RsDetails("branch_no").value = val(.TextMatrix(i, .ColIndex("branch_no")))
            RsDetails("ContNoteSerial1").value = (.TextMatrix(i, .ColIndex("ContNoteSerial1")))
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
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
            RsDetails("NoteDate").value = IIf(.TextMatrix(i, .ColIndex("NoteDate")) = "", Null, .TextMatrix(i, .ColIndex("NoteDate")))
            RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
            RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            RsDetails("TypTrans").value = 1
            RsDetails.update

      End If
      
                  If val(val(.TextMatrix(i, .ColIndex("NoteType")))) = -1 Then
                                    If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 And .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                                        StrSQL = "Update Notes Set  TotalPayed=1 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID2"))) & ""
                                        Cn.Execute StrSQL, , adExecuteNoRecords
                                     Else
                                         StrSQL = "Update Notes Set  TotalPayed=0 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID2"))) & ""
                                        Cn.Execute StrSQL, , adExecuteNoRecords
                                    End If
            Else
                                If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 And .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                                    StrSQL = "Update TblExpensesDet Set  TotalPayed=1 Where ID=" & val(.TextMatrix(i, .ColIndex("NoteID2"))) & ""
                                    Cn.Execute StrSQL, , adExecuteNoRecords
                                 Else
                                     StrSQL = "Update TblExpensesDet Set  TotalPayed=0 Where ID=" & val(.TextMatrix(i, .ColIndex("NoteID2"))) & ""
                                    Cn.Execute StrSQL, , adExecuteNoRecords
                                End If
          End If
          
    Next i
End With
    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     * from dbo.TblOwnerPayment202 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid3
    For i = .FixedRows To .rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID").value = val(XPTxtID.text)
            RsDetails("NoteType").value = val(.TextMatrix(i, .ColIndex("NoteType")))
            RsDetails("NoteID2").value = val(.TextMatrix(i, .ColIndex("NoteID2")))
            RsDetails("UnitNo").value = val(.TextMatrix(i, .ColIndex("UnitNo")))
            RsDetails("value").value = val(.TextMatrix(i, .ColIndex("value")))
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("NoteID3").value = val(.TextMatrix(i, .ColIndex("NoteID3")))
            RsDetails("TypTrans").value = 1
            RsDetails.update
        End If
    Next i
End With
End Function
Public Sub RetriveOwnerPaymentData()
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

    Set RsDetails = New ADODB.Recordset
  StrSQL = " SELECT     dbo.TblNotesOwnerPayment.ID, dbo.TblNotesOwnerPayment.NoteID, dbo.TblNotesOwnerPayment.PaymentID, dbo.TblNotesOwnerPayment.[value], "
  StrSQL = StrSQL & "                    dbo.TblNotesOwnerPayment.PayedValue, dbo.TblNotesOwnerPayment.RecDateH, dbo.TblNotesOwnerPayment.RecDate, dbo.TblNotesOwnerPayment.AllowDateH,"
  StrSQL = StrSQL & "                    dbo.TblNotesOwnerPayment.AllowDate, dbo.TblNotesOwnerPayment.TransPayedValue, dbo.TblNotesOwnerPayment.NetValue,"
  StrSQL = StrSQL & "                    dbo.TblNotesOwnerPayment.RemainingValue, dbo.TblAqrOwin.PaymentNo, dbo.TblNotesOwnerPayment.BranchId, dbo.TblBranchesData.branch_name,"
  StrSQL = StrSQL & "                    dbo.TblBranchesData.branch_namee , dbo.TblNotesOwnerPayment.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname"
  StrSQL = StrSQL & "  FROM         dbo.TblNotesOwnerPayment LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblAqrOwin ON dbo.TblNotesOwnerPayment.PaymentID = dbo.TblAqrOwin.ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblAqar ON dbo.TblNotesOwnerPayment.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.TblNotesOwnerPayment.BranchId = dbo.TblBranchesData.branch_id"
  StrSQL = StrSQL & "     Where (dbo.TblNotesOwnerPayment.NoteID = " & val(XPTxtID.text) & ")"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid1
    .Clear flexClearScrollable, flexClearEverything
    .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To RsDetails.RecordCount
        .TextMatrix(i, .ColIndex("Ser")) = i

            .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(RsDetails("BranchId").value), 0, RsDetails("BranchId").value)
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
            Else
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_namee").value), 0, RsDetails("branch_namee").value)
            End If
            .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(RsDetails("PaymentID").value), 0, RsDetails("PaymentID").value)
            .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDetails("value").value), 0, RsDetails("value").value)
            .TextMatrix(i, .ColIndex("PayedValue")) = IIf(IsNull(RsDetails("PayedValue").value), 0, RsDetails("PayedValue").value)
            .TextMatrix(i, .ColIndex("RecDateH")) = IIf(IsNull(RsDetails("RecDateH").value), "", RsDetails("RecDateH").value)
            .TextMatrix(i, .ColIndex("RecDate")) = IIf(IsNull(RsDetails("RecDate").value), "", RsDetails("RecDate").value)
            .TextMatrix(i, .ColIndex("AllowDateH")) = IIf(IsNull(RsDetails("AllowDateH").value), "", RsDetails("AllowDateH").value)
            .TextMatrix(i, .ColIndex("AllowDate")) = IIf(IsNull(RsDetails("AllowDate").value), "", RsDetails("AllowDate").value)
            .TextMatrix(i, .ColIndex("TransPayedValue")) = IIf(IsNull(RsDetails("TransPayedValue").value), 0, RsDetails("TransPayedValue").value)
            .TextMatrix(i, .ColIndex("RemainingValue")) = IIf(IsNull(RsDetails("RemainingValue").value), 0, RsDetails("RemainingValue").value)
            .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(RsDetails("NetValue").value), 0, RsDetails("NetValue").value)
            .TextMatrix(i, .ColIndex("PaymentNo")) = IIf(IsNull(RsDetails("PaymentNo").value), "", RsDetails("PaymentNo").value)
            .TextMatrix(i, .ColIndex("Aqarid")) = IIf(IsNull(RsDetails("Aqarid").value), 0, RsDetails("Aqarid").value)
            .TextMatrix(i, .ColIndex("aqarNo")) = IIf(IsNull(RsDetails("aqarNo").value), "", RsDetails("aqarNo").value)
            .TextMatrix(i, .ColIndex("aqarname")) = IIf(IsNull(RsDetails("aqarname").value), "", RsDetails("aqarname").value)
            '.TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(CDate(RsDetails("NoteDate").value))
            .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
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
Public Sub RetriveOwnerPaymentData202()
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

    Set RsDetails = New ADODB.Recordset
  StrSQL = " SELECT     dbo.TblNotesOwnerPayment202.ID,dbo.TblNotesOwnerPayment202.NoteType, dbo.TblNotesOwnerPayment202.NoteID, dbo.TblNotesOwnerPayment202.NoteID2, dbo.TblNotesOwnerPayment202.CusID, "
  StrSQL = StrSQL & "                     dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblNotesOwnerPayment202.Aqarid, dbo.TblAqar.aqarNo,"
  StrSQL = StrSQL & "                    dbo.TblAqar.aqarname, dbo.TblNotesOwnerPayment202.UnitNo, dbo.TblAqarDetai.unitno AS UntName, dbo.TblNotesOwnerPayment202.branch_no,"
  StrSQL = StrSQL & "                    dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblNotesOwnerPayment202.TypTrans, dbo.TblNotesOwnerPayment202.Remarks,"
  StrSQL = StrSQL & "                    dbo.TblNotesOwnerPayment202.ContNoteSerial1 , dbo.TblNotesOwnerPayment202.NoteDate, dbo.TblNotesOwnerPayment202.[value],"
  StrSQL = StrSQL & "                    dbo.TblNotesOwnerPayment202.PayedValue, dbo.TblNotesOwnerPayment202.RemainingValue, dbo.TblNotesOwnerPayment202.NetValue,"
  StrSQL = StrSQL & "                    dbo.TblNotesOwnerPayment202.TransPayedValue"
  StrSQL = StrSQL & "         FROM         dbo.TblNotesOwnerPayment202 LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.TblNotesOwnerPayment202.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblAqarDetai ON dbo.TblNotesOwnerPayment202.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblAqar ON dbo.TblNotesOwnerPayment202.Aqarid = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblCustemers ON dbo.TblNotesOwnerPayment202.CusID = dbo.TblCustemers.CusID"
  StrSQL = StrSQL & "  Where (dbo.TblNotesOwnerPayment202.TypTrans = 0) And (dbo.TblNotesOwnerPayment202.NoteID = " & val(XPTxtID.text) & ")"
  
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid2
    .Clear flexClearScrollable, flexClearEverything
    .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount
Frame12(1).Visible = True
        For i = .FixedRows To RsDetails.RecordCount
        .TextMatrix(i, .ColIndex("Ser")) = i

            .TextMatrix(i, .ColIndex("NoteType")) = IIf(IsNull(RsDetails("NoteType").value), 0, RsDetails("NoteType").value)
            If val(.TextMatrix(i, .ColIndex("NoteType"))) = 9 Then
            .TextMatrix(i, .ColIndex("NotesTypeName")) = "⁄—»Ê‰"
            Else
            .TextMatrix(i, .ColIndex("NotesTypeName")) = "„ﬁ»Ê÷« "
            End If
            .TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(RsDetails("branch_no").value), 0, RsDetails("branch_no").value)
            .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(RsDetails("CusID").value), 0, RsDetails("CusID").value)
            .TextMatrix(i, .ColIndex("Aqarid")) = IIf(IsNull(RsDetails("Aqarid").value), "", RsDetails("Aqarid").value)
            .TextMatrix(i, .ColIndex("NoteID2")) = IIf(IsNull(RsDetails("NoteID2").value), 0, RsDetails("NoteID2").value)
.TextMatrix(i, .ColIndex("InsuranceValue")) = GetInsuranceValue(val(.TextMatrix(i, .ColIndex("NoteID2"))))

            .TextMatrix(i, .ColIndex("aqarname")) = IIf(IsNull(RsDetails("aqarname").value), "", RsDetails("aqarname").value)
            .TextMatrix(i, .ColIndex("UnitNo")) = IIf(IsNull(RsDetails("UnitNo").value), "", RsDetails("UnitNo").value)
            .TextMatrix(i, .ColIndex("UntName")) = IIf(IsNull(RsDetails("UntName").value), "", RsDetails("UntName").value)
            .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks").value), "", RsDetails("Remarks").value)
            .TextMatrix(i, .ColIndex("ContNoteSerial1")) = IIf(IsNull(RsDetails("ContNoteSerial1").value), "", RsDetails("ContNoteSerial1").value)
            .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(RsDetails("NoteDate").value), "", RsDetails("NoteDate").value)
            .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDetails("value").value), 0, RsDetails("value").value)
            .TextMatrix(i, .ColIndex("PayedValue")) = IIf(IsNull(RsDetails("PayedValue").value), 0, RsDetails("PayedValue").value)
            .TextMatrix(i, .ColIndex("RemainingValue")) = IIf(IsNull(RsDetails("RemainingValue").value), 0, RsDetails("RemainingValue").value)
            .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(RsDetails("NetValue").value), 0, RsDetails("NetValue").value)
            .TextMatrix(i, .ColIndex("TransPayedValue")) = IIf(IsNull(RsDetails("TransPayedValue").value), 0, RsDetails("TransPayedValue").value)
            If SystemOptions.UserInterface = ArabicInterface Then
             .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsDetails("CusName").value), "", RsDetails("CusName").value)
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
            Else
            .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsDetails("CusNamee").value), "", RsDetails("CusNamee").value)
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_namee").value), 0, RsDetails("branch_namee").value)
            End If
            .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
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
Public Sub RetriveOwnerPaymentData203()
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

    Set RsDetails = New ADODB.Recordset
    
  StrSQL = " SELECT     dbo.TblNotesOwnerPayment202.ID, dbo.TblNotesOwnerPayment202.NoteType, dbo.TblNotesOwnerPayment202.NoteID,dbo.TblNotesOwnerPayment202.NoteID3, dbo.TblNotesOwnerPayment202.NoteID2, dbo.TblNotesOwnerPayment202.Aqarid, "
  StrSQL = StrSQL & "                    dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblNotesOwnerPayment202.UnitNo, dbo.TblAqarDetai.unitno AS UntName,"
  StrSQL = StrSQL & "                    dbo.TblNotesOwnerPayment202.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblNotesOwnerPayment202.TypTrans,"
  StrSQL = StrSQL & "                    dbo.TblNotesOwnerPayment202.Remarks, dbo.TblNotesOwnerPayment202.ContNoteSerial1, dbo.TblNotesOwnerPayment202.NoteDate,"
  StrSQL = StrSQL & "                    dbo.TblNotesOwnerPayment202.[value], dbo.TblNotesOwnerPayment202.PayedValue, dbo.TblNotesOwnerPayment202.RemainingValue,"
  StrSQL = StrSQL & "                    dbo.TblNotesOwnerPayment202.netvalue , dbo.TblNotesOwnerPayment202.TransPayedValue"
  StrSQL = StrSQL & "      FROM         dbo.TblNotesOwnerPayment202 LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.TblNotesOwnerPayment202.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblAqarDetai ON dbo.TblNotesOwnerPayment202.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblAqar ON dbo.TblNotesOwnerPayment202.Aqarid = dbo.TblAqar.Aqarid"
  StrSQL = StrSQL & "  Where (dbo.TblNotesOwnerPayment202.TypTrans = 1) And (dbo.TblNotesOwnerPayment202.NoteID = " & val(XPTxtID.text) & ")"
  
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid3
    .Clear flexClearScrollable, flexClearEverything
    .rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        .rows = .FixedRows + RsDetails.RecordCount
Frame12(1).Visible = True
        For i = .FixedRows To RsDetails.RecordCount
        .TextMatrix(i, .ColIndex("Ser")) = i
         .TextMatrix(i, .ColIndex("NoteType")) = IIf(IsNull(RsDetails("NoteType").value), 0, RsDetails("NoteType").value)
            If val(.TextMatrix(i, .ColIndex("NoteType"))) = -1 Then
            .TextMatrix(i, .ColIndex("NotesTypeName")) = " ’›Ì…"
            Else
            .TextMatrix(i, .ColIndex("NotesTypeName")) = "„’—Ê›« "
            End If
            .TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(RsDetails("branch_no").value), 0, RsDetails("branch_no").value)
            
            .TextMatrix(i, .ColIndex("Aqarid")) = IIf(IsNull(RsDetails("Aqarid").value), "", RsDetails("Aqarid").value)
            .TextMatrix(i, .ColIndex("NoteID2")) = IIf(IsNull(RsDetails("NoteID2").value), 0, RsDetails("NoteID2").value)
            .TextMatrix(i, .ColIndex("NoteID3")) = IIf(IsNull(RsDetails("NoteID3").value), 0, RsDetails("NoteID3").value)
            .TextMatrix(i, .ColIndex("aqarname")) = IIf(IsNull(RsDetails("aqarname").value), "", RsDetails("aqarname").value)
            .TextMatrix(i, .ColIndex("UnitNo")) = IIf(IsNull(RsDetails("UnitNo").value), "", RsDetails("UnitNo").value)
            .TextMatrix(i, .ColIndex("UntName")) = IIf(IsNull(RsDetails("UntName").value), "", RsDetails("UntName").value)
            .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks").value), "", RsDetails("Remarks").value)
            .TextMatrix(i, .ColIndex("ContNoteSerial1")) = IIf(IsNull(RsDetails("ContNoteSerial1").value), "", RsDetails("ContNoteSerial1").value)
            .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(RsDetails("NoteDate").value), "", RsDetails("NoteDate").value)
            .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDetails("value").value), 0, RsDetails("value").value)
            .TextMatrix(i, .ColIndex("PayedValue")) = IIf(IsNull(RsDetails("PayedValue").value), 0, RsDetails("PayedValue").value)
            .TextMatrix(i, .ColIndex("RemainingValue")) = IIf(IsNull(RsDetails("RemainingValue").value), 0, RsDetails("RemainingValue").value)
            .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(RsDetails("NetValue").value), 0, RsDetails("NetValue").value)
            .TextMatrix(i, .ColIndex("TransPayedValue")) = IIf(IsNull(RsDetails("TransPayedValue").value), 0, RsDetails("TransPayedValue").value)
            If SystemOptions.UserInterface = ArabicInterface Then
             
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
            Else
            
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_namee").value), 0, RsDetails("branch_namee").value)
            End If
            .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
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
Sub RelineBuy()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid1
        For i = .FixedRows To .rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           Sm = Sm + val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           End If
           Next i
  
    End With
  Label27(1).Caption = Sm
End Sub

Function GeteOwnerPayed(Optional PaymentID As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS SumValue"
sql = sql & " From dbo.TblOwnerPayment"
sql = sql & " Where (PaymentID = " & PaymentID & ")"
sql = sql & " GROUP BY PaymentID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteOwnerPayed = IIf(IsNull(Rs8("SumValue").value), 0, Rs8("SumValue").value)
Else
GeteOwnerPayed = 0
End If
End Function
Function GeteOwnerPayedOPening(Optional PaymentID As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PreBalaTransPyed) AS SumValue"
sql = sql & " from Notes where NoteType=5 and NoteID<>" & val(XPTxtID.text) & ""
sql = sql & " and   IqarID2=" & val(Me.DcbIqara2.BoundText)
'rs("IqarID2").value = val(Me.DcbIqara2.BoundText)

'salah mnh llah
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteOwnerPayedOPening = IIf(IsNull(Rs8("SumValue").value), 0, Rs8("SumValue").value)
Else
GeteOwnerPayedOPening = 0
End If
End Function
Function GeteOwnerPayed202(Optional PaymentID As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS SumValue"
sql = sql & " From dbo.TblOwnerPayment202"
sql = sql & " Where (NoteID2 = " & PaymentID & ") and TypTrans=0 "
sql = sql & " GROUP BY NoteID2"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteOwnerPayed202 = IIf(IsNull(Rs8("SumValue").value), 0, Rs8("SumValue").value)
Else
GeteOwnerPayed202 = 0
End If
End Function
Function GeteOwnerPayed203(Optional PaymentID As Double = 0) As Double
If PaymentID = 0 Then
GeteOwnerPayed203 = 0
Exit Function
End If
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS SumValue"
sql = sql & " From dbo.TblOwnerPayment202"
sql = sql & " Where (NoteID2 = " & PaymentID & ") and TypTrans=1 and (UnitNo=0 or UnitNo is null) "
sql = sql & " GROUP BY NoteID2"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteOwnerPayed203 = IIf(IsNull(Rs8("SumValue").value), 0, Rs8("SumValue").value)
Else
GeteOwnerPayed203 = 0
End If
End Function
Function GeteOwnerPayed204(Optional PaymentID As Double = 0) As Double
If PaymentID = 0 Then
GeteOwnerPayed204 = 0
Exit Function
End If
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS SumValue"
sql = sql & " From dbo.TblOwnerPayment202"
sql = sql & " Where (NoteID3 = " & PaymentID & ") and TypTrans=1 and not(UnitNo is null) "
sql = sql & " GROUP BY NoteID2"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteOwnerPayed204 = IIf(IsNull(Rs8("SumValue").value), 0, Rs8("SumValue").value)
Else
GeteOwnerPayed204 = 0
End If
End Function
Private Sub CboPayMentType_Change()

    If Me.TxtModFlg.text = "E" Then
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        Me.DcboBox.text = ""
    End If

    Frame4.Visible = False

    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(16).Caption = "—ﬁ„ «·‘Ìﬂ "
        lbl(17).Caption = " «—ÌŒ «·«” Õﬁ«ﬁ"
    
    Else
        lbl(16).Caption = "Cheque No"
        lbl(17).Caption = "Due Date"
    End If

    If Me.CboPayMentType.ListIndex = 0 Then
        Me.lbl(9).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
    ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 3 Then
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
    ElseIf Me.CboPayMentType.ListIndex = 2 Then
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
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
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
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    Dim auto_sanad_no As String
    Dim NO As Integer
    auto_sanad_no = ""
    departement_name = 1
 
    connection_string = Cn.ConnectionString
    numbering.ConnectionString = connection_string
    numbering.CommandType = adCmdText
    numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=4"
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
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=5 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
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
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=5 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 4, 2)
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
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=5 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4)
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
                    auto_sanad_no = mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & start_at

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

Public Sub Cmd_Click(index As Integer)
    Dim cNoteReport As ClsNotesReports
    Dim Msg As String
   ' On Error GoTo ErrTrap

    

    Select Case index

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
            TxtModFlg.text = "N"
            ' XPTxtID.text = CStr(new_id("Notes", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=5"))
             TotalPayments.text = 0
TxtTotalPayedOpBalance.text = 0
txtPercent.text = 0
TxtValuExpenses.text = 0
TxtNetPayments.text = 0
TxtOfficeValue.text = 0
TxtOfficeValueDiscAdd.text = 0
TxtOfficeValueNet.text = 0
TxtNetValue.text = 0
            Me.DCboUserName.BoundText = user_id
            VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid3.rows = 1
            VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.rows = 1
  
            '      XPDtbTrans.SetFocus
            Text1.text = setfoxy
            Me.dcBranch.BoundText = Current_branch
            Option1.value = False
            Option2.value = False
            Option3.value = False
            Txt_DateHigri.value = ToHijriDate(Date)
XPDtbTrans.SetFocus
Option3.value = True
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
        '    Me.DCboUserName.BoundText = user_id
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

                With Fg
 
                    Me.LblTotalV.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("PartValue"), .rows - 1, .ColIndex("PartValue"))
              
                End With
    
                If Round(LblTotalV.Caption, 2) <> Round(val(XPTxtVal.text), 2) Then
    
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
     
Dim RentAccount As String
                If val(TxtOfficeValue.text) > 0 And SystemOptions.NoCreatJLInRentContract = True And val(DCboCashType.ListIndex) = 8 Then
                GetValueAddedAccount XPDtbTrans.value, , RentAccount, 1, 21
                If RentAccount = "" Then
                MsgBox ("Ì—ÃÏ ÷»ÿ «⁄œ«œ  «·ﬁÌ„… «·„÷«›… ··„»Ì⁄« ")
                Exit Sub
                End If
                   RentAccount = get_account_code_branch(207, dcBranch.BoundText)
                   If RentAccount = "" Or RentAccount = "NO account" Then
                   MsgBox ("Ì—ÃÏ «Œ Ì«— Õ”«» «·⁄„Ê·« ")
                   Exit Sub
                   End If
               End If
    
            SaveData
        SendMessage (0)
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

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 2020
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
              '  print_report Me.TxtNoteSerial.text, Me.TxtCustCode.text
        
                '     Set cNoteReport = New ClsNotesReports
                '     cNoteReport.PrintReceipt Val(Me.XPTxtID.text), WindowTarget
                '     Set cNoteReport = Nothing
                print_report
                SendMessage (1)
            End If

        Case 9
   
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

  '          ShowGL_cc Me.TxtNoteSerial.Text, , 200
   ShowGL_cc Me.TxtNoteSerial.text, , 200, val(XPTxtID.text) ', txtTotalWithVat
   
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
 
            Me.DCboUserName.BoundText = user_id
              'Me.DcBranch.BoundText = Current_branch
     TxtNoteSerial.text = ""
     TxtNoteSerial1.text = ""
     
    called = False
    Case 13
    print_report2
   Case 14
    print_report3
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Function CheckDate() As Boolean
    Dim StrTemp As String
    Dim Msg  As String

    If year(Date) > val(Me.CboYear.text) Then ' ⁄«„ „÷Ï
        Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ ÕÌÀ «‰Â ﬁ»·  «—ÌŒ «·ÌÊ„ ÊÂ–« ·« ÌÃÊ“...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CheckDate = False
        Exit Function
    ElseIf year(Date) = val(Me.CboYear.text) Then '‰›” «·⁄«„

        If Month(Date) > val(Me.CmbMonth.ListIndex + 1) Then
            'Msg = "«· «—ÌŒ «·„Õœœ €Ì— ’ÕÌÕ...!!!"
            'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            'CheckDate = False
            'Exit Function
        End If
    End If

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
End Sub

Private Sub CalCulateParts()
    Dim i As Integer
    Dim IntPartCounts As Integer
    Dim SngPartValue As Double
    Dim m_FirstDate As Date

    If CheckPartCal = False Then
        Exit Sub
    End If

    If CheckDate = False Then
        Exit Sub
    End If

    SngPartValue = val(Me.XPTxtVal.text) / val(Me.TxtPaymentCounts.text)
    IntPartCounts = val(Me.TxtPaymentCounts.text)
    m_FirstDate = CDate(val(Me.CboYear.text) & "-" & Me.CmbMonth.ListIndex + 1 & "-01")

    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .rows = .FixedRows + IntPartCounts
        .RowHeightMin = 300

        For i = 1 To IntPartCounts
            .TextMatrix(i, .ColIndex("PartNO")) = i
            .TextMatrix(i, .ColIndex("PartValue")) = Round(SngPartValue, 2)
            .TextMatrix(i, .ColIndex("PartDate")) = DisplayDate(DateAdd("m", i - 1, m_FirstDate))
        Next i
    
    End With

End Sub

'Public Function print_report(Optional NoteSerial As String, Optional Custcode As String)
'
'    Dim MySQL As String
'    Dim RsData As New ADODB.Recordset
'    Dim xApp As New CRAXDRT.Application
'    Dim xReport As CRAXDRT.Report
'    Dim CViewer As ClsReportViewer
'    Dim StrReportTitle As String
'    Dim StrFileName As String
'    Dim Msg As String

'    MySQL = "Select * From EXPENSES_ORDER2  where NoteSerial='" & NoteSerial & "'"
 
'    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
'    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
'    'End If
'    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
'    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
'    'End If
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        StrFileName = App.path & "\Reports\" & "Expenses_order3.rpt"
'    Else
'        StrFileName = App.path & "\Reports\" & "Expenses_order3_Eng.rpt"
'    End If
'
'    If Dir(StrFileName) = "" Then
'        'GetMsgs 139, vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    Set RsData = New ADODB.Recordset
'    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
'
'    Screen.MousePointer = vbArrowHourglass
'    Set xReport = xApp.OpenReport(StrFileName)
'    xReport.Database.SetDataSource RsData
'
'    Dim cCompanyInfo As New ClsCompanyInfo
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
'        xReport.ParameterFields(5).AddCurrentValue DcboDebitSide.text  'RPTCompany_Name_Arabic
'        xReport.ParameterFields(6).AddCurrentValue Custcode
'        xReport.ParameterFields(7).AddCurrentValue DcboCreditSide.text  'RPTCompany_Name_Arabic   xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.text  'RPTCompany_Name_Arabic
''
'        'CustCode
'        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
'        StrReportTitle = "" '& StrAccountName
'        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
'        'End If
'        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
'        'End If
'    Else
'
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
'        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
'        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
'        xReport.ParameterFields(5).AddCurrentValue DcboDebitSide.text
'        xReport.ParameterFields(6).AddCurrentValue Custcode
'        StrReportTitle = ""
'        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
'        'End If
'        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
'        'End If
'    End If
'
'    xReport.ParameterFields(3).AddCurrentValue user_name
'    xReport.reporttitle = StrReportTitle
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.title
'    xReport.ReportAuthor = App.title
'    Set CViewer = New ClsReportViewer
'    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
'
'    RsData.Close
'    Set RsData = Nothing
'    Screen.MousePointer = vbDefault
'
'End Function

Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments Me.TxtNoteSerial1, "0712201401"

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
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
Sub DeleteOwner202()
Dim i As Integer
Dim StrSQL As String
With VSFlexGrid2
 For i = .FixedRows To .rows - 1
 If val(.TextMatrix(i, .ColIndex("NoteID2"))) <> 0 Then
      StrSQL = "Update Notes Set  TotalPayed=0 Where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID2"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
     End If
     Next i
 End With
End Sub
Sub DeleteOwner203()
Dim i As Integer
Dim StrSQL As String
With VSFlexGrid3
 For i = .FixedRows To .rows - 1
 If val(.TextMatrix(i, .ColIndex("NoteID2"))) <> 0 Then
      StrSQL = "Update TblExpensesDet Set  TotalPayed=0 Where ID=" & val(.TextMatrix(i, .ColIndex("NoteID2"))) & ""
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
 If val(.TextMatrix(i, .ColIndex("PaymentID"))) <> 0 Then
      StrSQL = "Update TblAqrOwin Set  TotalPayed=0 Where ID=" & val(.TextMatrix(i, .ColIndex("PaymentID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
     End If
     Next i
 End With
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command10_Click(index As Integer)
Dim i As Integer
Dim StrSQL As String
If index = 0 Then
If Me.TxtModFlg.text = "E" Then
DeleteBillBuy
DeleteOwner203

VSFlexGrid1.Enabled = True
        Check1.Enabled = True
      StrSQL = "Delete From TblNotesOwnerPayment Where NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblOwnerPayment Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.text)
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
End Sub

Private Sub Command11_Click()
Dim i As Integer
Dim StrSQL As String
If Me.TxtModFlg.text = "E" Then
DeleteOwner202
VSFlexGrid2.Enabled = True
        Check1.Enabled = True
      StrSQL = "Delete From TblNotesOwnerPayment202 Where  NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblOwnerPayment202 Where   NoteID=" & val(Me.XPTxtID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords

VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
VSFlexGrid2.rows = 1
DeleteOwner203
VSFlexGrid3.Enabled = True
        Check1.Enabled = True
VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
VSFlexGrid3.rows = 1
FlgBillBuy = True
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
            With Me.VSFlexGrid3

            For i = .FixedRows To .rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i


        End With
End If


End Sub



Public Sub DBCboClientName_Change()

    On Error Resume Next
    TxtCustCode.text = ""
    Dim DefaultSalesPersonId As Integer
    Dim Fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, Fullcode
    TxtCustCode.text = Fullcode
 
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If DCboCashType.ListIndex = 3 Or DCboCashType.ListIndex = 4 Or DCboCashType.ListIndex = 5 Then
            Me.DcboDebitSide.BoundText = DBCboClientName.BoundText
        Else
        If SystemOptions.NoCreatJLInRentContract = True And DCboCashType.ListIndex = 8 Then
           ' Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
          If SystemOptions.OpenAccountAqar = False Then
            Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Txtownerid), "Account_code")
         Else
            Me.DcboDebitSide.BoundText = GetAqarAcountCode(val(DcbIqara.BoundText))
         End If
         Else
            Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
         End If
        End If
        Dim Dcombos As ClsDataCombos
'    Dcombos.GetAccountingCodes Me.DcboDebitSide
'    Dcombos.GetAccountingCodes Me.DcboCreditSide
        Dim lblflag As Integer
        
        
                     If DCboCashType.ListIndex = 4 Then
            
                    If Option4.value = True Then
                    lblflag = 1
                   ElseIf Option5.value = True Then
                    lblflag = 0
            
                   ElseIf Option6.value = True Then
                    lblflag = 2
                  ElseIf Option7.value = True Then
                    lblflag = 3
                   End If
            
            
            
              GetEmployeeIDFromCode , , , Fullcode, , lblflag, DBCboClientName.BoundText, True
                   TxtCustCode.text = Fullcode
              End If
  



    End If

End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If DCboCashType.ListIndex = 0 Then
        If KeyCode = vbKeyF3 Then
            FrmCustemerSearch.show vbModal
            FrmCustemerSearch.SearchType = 1915
        End If

    ElseIf DCboCashType.ListIndex = 1 Then

        If KeyCode = vbKeyF3 Then
            FrmCompanySearch.show vbModal
              FrmCompanySearch.lblSearchtype.Caption = 19152
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
            Account_search.case_id = 1300
            
        End If



    End If

End Sub

Private Sub DcbIqara_Change()
DcbIqara_Click (0)
End Sub

Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then: Exit Sub
Dim EmpCode  As String
Dim ownerid As Double
GetIqarCode , , DcbIqara.BoundText, EmpCode, ownerid
    Me.TxtSearch.text = EmpCode
   
End Sub

Public Sub DcbIqara2_Change()
DcbIqara2_Click (0)
DBCboClientName.Enabled = False
End Sub

Public Sub DcbIqara2_Click(Area As Integer)
      If val(DcbIqara2.BoundText) = 0 Then: Exit Sub
Dim EmpCode  As String
Dim ownerid As Double
GetIqarCode , , DcbIqara2.BoundText, EmpCode, ownerid
    Me.TxtSearch2.text = EmpCode
    DBCboClientName.BoundText = ownerid
    DcbIqara.BoundText = DcbIqara2.BoundText
    Me.TxtSearch.text = Me.TxtSearch2.text
    If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
    With VSFlexGrid3
      .Clear flexClearScrollable, flexClearEverything
      .rows = 1
    End With
  With VSFlexGrid2
    .Clear flexClearScrollable, flexClearEverything
    .rows = 1
   End With
End If
DBCboClientName_Change
End Sub

Private Sub DcboBankName_Click(Area As Integer)
    On Error Resume Next

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

Private Sub DCboCashType_Change()
    Frame2.Visible = False
    Option4.value = False
    Option5.value = False
    Option6.value = False
    Option7.value = False
    Frame12(0).Visible = False
    ALLButton4.Visible = False
    DBCboClientName = ""
    Frame10.Visible = False
    Frame12(1).Visible = False
    ISButton1.Visible = False
    DBCboClientName.Enabled = True
    TxtCustCode.Enabled = True
    Dim StrSQL As String
    Dim intDef As Integer
    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String
    ISButton3.Visible = False
    txtTotalinsuranceS.Visible = False
    lbl(49).Visible = False
Frame5.Visible = False
Frame6.Visible = False
    If SystemOptions.UserInterface = EnglishInterface Then
        lbl(3).Caption = "Name"
    Else
        lbl(3).Caption = "«·«”„"
    End If
        
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "E" Then

        With Fg
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows + 1
    
        End With

        Fra(2).Visible = False
    End If

    Select Case DCboCashType.ListIndex

        Case 0
            Set Dcombos = New ClsDataCombos
            Dcombos.GetCustomersSuppliers 156, Me.DBCboClientName, True
            ChkTrans.Visible = True
            Fra(0).Visible = True

        Case 1
            Set Dcombos = New ClsDataCombos
            Dcombos.GetCustomersSuppliers 257, Me.DBCboClientName, True
            ChkTrans.Visible = True
            Fra(0).Visible = True

        Case 2
            Set Dcombos = New ClsDataCombos
            Dcombos.GetPersons Me.DBCboClientName
            ChkTrans.Visible = False
            Fra(0).Visible = False

        Case 3
            Fra(0).Visible = True

            If SystemOptions.UserInterface = EnglishInterface Then
                lbl(3).Caption = "Project"
            Else
                lbl(3).Caption = "«·„‘—Ê⁄"
            End If

            Set Dcombos = New ClsDataCombos
            Dcombos.ClearMyDataCombo Me.DBCboClientName
            If SystemOptions.UserInterface = ArabicInterface Then
                    My_SQL = "  select expanses_account,Project_name from projects where not(expanses_account is null)  order by Project_name " '  where  Account_code like'" & Account_Code_dynamic & "%' and last_account=1"
            Else
                    My_SQL = "  select expanses_account,Project_namee from projects where not(expanses_account is null)  order by Project_name " '  where  Account_code like'" & Account_Code_dynamic & "%' and last_account=1"
            End If
            
            fill_combo Me.DBCboClientName, My_SQL
 
        Case 4
            Frame2.Visible = True
            Frame2.Enabled = True

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
       
            Set Dcombos = New ClsDataCombos
            Dcombos.ClearMyDataCombo Me.DBCboClientName
         
            fill_combo Me.DBCboClientName, My_SQL
      
        Case 5
            Set Dcombos = New ClsDataCombos
            Dcombos.ClearMyDataCombo Me.DBCboClientName
If SystemOptions.UserInterface = ArabicInterface Then
            My_SQL = "  select Account_Code,Account_Name from ACCOUNTS where last_account=1"
Else
         My_SQL = "  select Account_Code,Account_Nameeng from ACCOUNTS where last_account=1"
End If
            fill_combo Me.DBCboClientName, My_SQL
       
            '   My_SQL = "  select Account_Code,BoxName from TblBoxesData where Type=1"
            '  fill_combo Me.DBCboClientName, My_SQL
   

        Case 6
         txtTotalinsuranceS.Visible = True
    lbl(49).Visible = True
        
                  ISButton3.Visible = True
'Frame5.Visible = True
  Set Dcombos = New ClsDataCombos
            Dcombos.GetCustomersSuppliers 56, Me.DBCboClientName, True
            DBCboClientName.Enabled = False
            TxtCustCode.Enabled = False
Case 7
ISButton1.Visible = True
Frame6.Visible = True
        
                 
'Frame5.Visible = True
  Set Dcombos = New ClsDataCombos
          '  Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
          '  DBCboClientName.Visible = False
            
            
         Dim Account_Code_dynamic As String
                    Account_Code_dynamic = get_account_code_branch(95, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created ", vbCritical
            End If

            GoTo ErrTrap
        ElseIf Account_Code_dynamic = "NO account" Then

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „œ›Ê⁄«  „ﬁœ„… ·ÕÃ“ «·ÊÕœ«  ", vbCritical
            Else
                MsgBox "   Insatllemts Revenu Not Deined in this Branch", vbCritical
            End If

            GoTo ErrTrap
         
   
    
        End If
             '    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboDebitSide.BoundText = Account_Code_dynamic
        
            
            
            
'Case 8
'   If SystemOptions.UserInterface = ArabicInterface Then
'            My_SQL = "  select Account_Code,BoxName from TblBoxesData where Type=0"
'    Else
'
'    My_SQL = "  select Account_Code,BoxNameE from TblBoxesData where Type=0"
'
'    End If
'     fill_combo Me.DBCboClientName, My_SQL
   
    Case 8
    Option2.value = True
    
    DBCboClientName.Enabled = False
  ' Dcombos.GetCustomersSuppliers 57, Me.DBCboClientName
    Frame10.Visible = True
  ALLButton4.Visible = True
  Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 57, Me.DBCboClientName, True
    If Me.TxtModFlg.text = "R" Or Me.TxtModFlg.text = "" Then
    'Frame12(0).Visible
    End If
            
    End Select

    cSearchDcbo.Refresh
    Set Dcombos = Nothing
    Exit Sub
ErrTrap:

End Sub

Private Sub DCboCashType_Click()
    DCboCashType_Change
End Sub

Private Sub ChangeLang()
    lbl(22).Caption = "Curr. Week"
    lbl(40).Caption = "Branch"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Frame1.Caption = "Options"
    Fra(2).Caption = "detalis"
    CmdAttach.Caption = "Attachments"

    
    Option3.Caption = "Adv. Payment"
    Option2.Caption = "Select Invoice"
    ALLButton3.Caption = "Select"
    lbl(37).Caption = "Order No :"
    lbl(22).Caption = "Current Week"
    lbl(35).Caption = "Adv. Pay."
    Label8.Caption = "General C.C."
    lbl(36).Caption = "General Des"
    Cmd(9).Caption = "GL Print"
    Cmd(10).Caption = "Cheque Print"
    Frame2.Caption = "Employee"
    Option4.Caption = "Salary"
    Option5.Caption = "Advanced"
    Option6.Caption = "Alloc"
    Option7.Caption = "Adv. Paayment"

    ALLButton1.Caption = "Installment view"
    ALLButton2.Caption = "debt Voucher"
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
    CmdHelp.Caption = "Help"
    DCboCashType.Clear
    DCboCashType.AddItem "To Customer"
    DCboCashType.AddItem "To Vendor"
    DCboCashType.AddItem "sub-contractor"
    DCboCashType.AddItem "To Project"
    DCboCashType.AddItem "To Employee"
    DCboCashType.AddItem "To Acc."
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
      
    End With

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

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub
 

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcbUnitNo_Change()
DcbUnitNo_Click (0)
End Sub

Private Sub DcbUnitNo_Click(Area As Integer)
Dim Typed As Integer
If Me.TxtModFlg.text <> "R" Then
GetIqarUnitData val(DcbUnitNo.BoundText), , , , , , , , , , , , , Typed
ComResid(Typed).value = True
End If
End Sub

Private Sub DcbUnitType_Change()
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
   Set Dcombos = New ClsDataCombos
  ' Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"

If val(DcbIqara.BoundText) > 0 Then
idd = val(DcbIqara.BoundText)

idd1 = val(DcbUnitType.BoundText)
If Me.TxtModFlg = "R" Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
Else
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
End If
End If
End Sub

Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 5
    End If

End Sub

Public Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.DateOpt = 1 Then
        Txt_DateHigri.Visible = True
    
    End If


If SystemOptions.SpecialVersion = True Then
Cmd(7).Visible = False
Fra(1).Visible = False
End If
   
   
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos

'    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
'    fill_combo Me.DcCostCenter, StrSQL

'    Dim Dcombos As ClsDataCombos
'Set Dcombos = New ClsDataCombos




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
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    'Resize_Form Me

    AddTip
    DCboCashType.AddItem "≈·Ï ⁄„Ì·"
    DCboCashType.AddItem "≈·Ï „Ê—œ"
    DCboCashType.AddItem "„ﬁ«Ê· »«ÿ‰"
    DCboCashType.AddItem "„‘—Ê⁄"
    DCboCashType.AddItem "„ÊŸ›"
    DCboCashType.AddItem "Õ”«»"
    DCboCashType.AddItem " ’›ÌÂ"
    DCboCashType.AddItem " ⁄—»Ê‰"
    DCboCashType.AddItem "«·„·«ﬂ"

    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetIqarUnit -2, 1, Me.DcbUnitNo
Dcombos.GetCostCenter DcCostCenter
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    With Me.CboPayMentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "ÕÊ«·Â »‰ﬂÌÂ"
        .AddItem "  ‘Ìﬂ „”œœ"
      
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
    Dcombos.GetIqar DcbIqara2
    Dcombos.GetIqar DcbIqara
    Dcombos.getAkarUnit Me.DcbUnitType

    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=5 and cashingtype<=8 "
       StrSQL = StrSQL & "     AND branch_no in(" & Current_branchSql & ")"
        StrSQL = StrSQL & " and  not (  (akarid is null )  and   (IqarID2 is null )  and   (NoteOrBonID is null ) )  "
        
      '  If SystemOptions.usertype <> UserAdminAll Then
      '  StrSQL = StrSQL & " AND   branch_no=" & Current_branch
   ' End If
    
    StrSQL = StrSQL & "order by NoteID "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    SetDtpickerDate XPDtbTrans
    SetDtpickerDate Me.DtpChequeDueDate
    ChkTrans.value = Unchecked
    ChkTrans_Click

    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"
    WriteInfo
    Dim My_SQL As String

    'My_SQL = "  select account_no,account_name from projects  where not (account_no is null)"
    My_SQL = "  select expanses_account,Project_name from projects where not(expanses_account is null)" '  where  Account_code like'" & Account_Code_dynamic & "%' and last_account=1"
    fill_combo dcproject, My_SQL

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
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

Private Sub FrmPriodDate_Change()
   If Me.TxtModFlg.text <> "R" Then
     
    FrmPriodDateH.value = ToHijriDate(FrmPriodDate.value)
    
End If
End Sub

Private Sub FrmPriodDateH_LostFocus()
 If Me.TxtModFlg.text <> "R" Then
             
             FrmPriodDate.value = ToGregorianDate(FrmPriodDateH.value)

               
        End If
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption



End Sub

Private Sub ISButton1_Click()
  Load FrmNotesSearch
           FrmNotesSearch.SearchType = 10
            FrmNotesSearch.show vbModal
End Sub

Private Sub ISButton3_Click()
FrmIqarWaiverSet.m_RetrunType = 7
 FrmIqarWaiverSet.show vbModal
End Sub

Private Sub Label29_Click()
Frame12(0).Visible = False
End Sub

Private Sub Label3_Click()
Frame12(1).Visible = False
End Sub

Private Sub Label4_Click()

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
                              x As Single, _
                              Y As Single)

    If SystemOptions.UserInterface = ArabicInterface Then
        LblLink.ToolTipText = "—’Ìœ «·ÿ—› «·„œÌ‰:" & WriteNo(Balance, 0, True)
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        LblLink.ToolTipText = "Depit Balance:" & WriteNo(Balance, 0, True)
    End If

End Sub

Private Sub optAdd_Click()
    RelineOwner2
End Sub

Private Sub optDisc_Click()
optAdd_Click
End Sub

Private Sub Option1_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

End Sub

Private Sub Option2_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
        ALLButton4.Enabled = True
    Else
        ALLButton3.Enabled = False
        ALLButton4.Enabled = False
    End If

End Sub

Private Sub Option3_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

End Sub

Private Sub Option4_Click()
    Dim My_SQL As String
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select Account_Code1,Emp_Name from TblEmployee where   (Account_Code1 <> N'""' AND NOT (Account_Code1 IS NULL)) "
Else
My_SQL = "  select Account_Code1,Emp_Namee from TblEmployee where   (Account_Code1 <> N'""' AND NOT (Account_Code1 IS NULL)) "
End If
    fill_combo Me.DBCboClientName, My_SQL

    If txt_general_des.text = "" And Me.TxtModFlg <> "R" Then
        txt_general_des.text = Option4.Caption
    End If

    Fra(2).Visible = False
End Sub

Private Sub Option5_Click()
    Dim My_SQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select Account_Code,Emp_Name from TblEmployee    where (Account_Code <> N'""' AND NOT (Account_Code IS NULL)) "
    Else
    My_SQL = "  select Account_Code,Emp_Namee from TblEmployee    where (Account_Code <> N'""' AND NOT (Account_Code IS NULL)) "
    End If
    fill_combo Me.DBCboClientName, My_SQL

    If txt_general_des.text = "" And Me.TxtModFlg <> "R" Then
        txt_general_des.text = Option5.Caption
    End If

    Fra(2).Visible = True
End Sub



Private Sub Option6_Click()
    Dim My_SQL As String
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select Account_Code2,Emp_Name from TblEmployee where   (Account_Code2 <> N'""' AND NOT (Account_Code2 IS NULL)) "
 Else
 My_SQL = "  select Account_Code2,Emp_Namee from TblEmployee where   (Account_Code2 <> N'""' AND NOT (Account_Code2 IS NULL)) "
 End If
 
    fill_combo Me.DBCboClientName, My_SQL

    If txt_general_des.text = "" And Me.TxtModFlg <> "R" Then
        txt_general_des.text = Option6.Caption
    End If

    Fra(2).Visible = False
End Sub

Private Sub Option7_Click()
    Dim My_SQL As String
If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = "  select Account_Code3,Emp_Name from TblEmployee  where (Account_Code3 <> N'""' AND NOT (Account_Code3 IS NULL)) "
Else
My_SQL = "  select Account_Code3,Emp_Namee from TblEmployee  where (Account_Code3 <> N'""' AND NOT (Account_Code3 IS NULL)) "
End If

    fill_combo Me.DBCboClientName, My_SQL

    If txt_general_des.text = "" And Me.TxtModFlg <> "R" Then
        txt_general_des.text = Option7.Caption
    End If

    Fra(2).Visible = False
End Sub

Private Sub ToPriodDate_Change()
   If Me.TxtModFlg.text <> "R" Then
     
    ToPriodDateH.value = ToHijriDate(ToPriodDate.value)
    
End If
End Sub

Private Sub ToPriodDateH_LostFocus()
 If Me.TxtModFlg.text <> "R" Then
             
             ToPriodDate.value = ToGregorianDate(ToPriodDateH.value)

               
        End If
End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
    
        Order_no_search.show
        Order_no_search.RetrunType = 2
    End If

End Sub

Private Sub TxtCommission_Change()
If Me.TxtModFlg <> "R" Then
txtTotal1.text = val(Txtcommission.text) + val(TxtWater.text) + val(txtinstrunce.text)
If val(XPTxtVal.text) >= (val(Txtcommission.text) - val(TxtCommissionOut.text)) Then
txtComisin.text = val(Txtcommission.text) - val(TxtCommissionOut.text)
Else
txtComisin.text = val(txtDiff.text)
End If
End If
End Sub

Private Sub TxtCommissionOut_Change()
If Me.TxtModFlg <> "R" Then
txtTotal1.text = val(Txtcommission.text) + val(TxtWater.text) + val(txtinstrunce.text)
If val(XPTxtVal.text) >= (val(Txtcommission.text) - val(TxtCommissionOut.text)) Then
txtComisin.text = val(Txtcommission.text) - val(TxtCommissionOut.text)
Else
txtComisin.text = val(txtDiff.text)
End If
End If
End Sub

Private Sub TxtCustCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer
Dim Typ As Integer
If val(DCboCashType.ListIndex) = 8 Then
Typ = 57
Else
Typ = DCboCashType.ListIndex + 1
End If
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtCustCode.text, Typ
        
        DBCboClientName.BoundText = CUSTID
    End If

Dim lblflag As Integer
 If DCboCashType.ListIndex = 4 Then


        If Option4.value = True Then
        lblflag = 1
       ElseIf Option5.value = True Then
        lblflag = 0
       
       ElseIf Option6.value = True Then
        lblflag = 2
      ElseIf Option7.value = True Then
        lblflag = 3
       End If

Dim Account_code As String

  GetEmployeeIDFromCode TxtCustCode.text, , , , , lblflag, Account_code
        DBCboClientName.BoundText = Account_code
 End If
        
End Sub

Private Sub txtinstrunce_Change()
If Me.TxtModFlg <> "R" Then
txtTotal1.text = val(Txtcommission.text) + val(TxtWater.text) + val(txtinstrunce.text)
End If
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            Frame2.Enabled = False

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Payments"
            Else
                Me.Caption = "«·„œ›Ê⁄« "
     
            End If

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

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Payments (Edit)"
            Else
                Me.Caption = "«·„œ›Ê⁄« (  ⁄œÌ· )"
        
            End If
    
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
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

Private Sub TxtNetValue_Change()
XPTxtVal.text = val(TxtNetValue)
End Sub

Private Sub TxtNotID_Change()
If Me.TxtModFlg.text <> "R" Then
If val(Me.TxtNotID.text) <> 0 Then
GetNotesInformation val(Me.TxtNotID.text)
End If
End If
End Sub

Private Sub TxtNotVal_Change()
txtDiff.text = val(Me.TxtNotVal.text) - val(XPTxtVal.text)
End Sub

Private Sub TxtOfficeValueDiscAdd_Change()
optAdd_Click
End Sub

Private Sub TxtPreBalaTransPyed_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If val(TxtPreBalaTransPyed.text) > val(TxtPreBalaRemain.text) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬁÌ„… «ﬂ»— „‰ «·„ »ﬁÌ"
Else
MsgBox "Paid value is greater than remaining"
End If
TxtPreBalaTransPyed.text = 0
TxtTotalPayedOpBalance.text = 0
Exit Sub
End If
RelineOwner2
End If

End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
Dim EmpID As Double

    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.text, EmpID
        DcbIqara.BoundText = EmpID
        DcbIqara_Click (0)
    End If
End Sub

Private Sub TxtSearch2_KeyPress(KeyAscii As Integer)
Dim EmpID As Double
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch2.text, EmpID
        DcbIqara2.BoundText = EmpID
        DcbIqara2_Click (0)
    End If
End Sub

Private Sub txttotal1_Change()
If Me.TxtModFlg <> "R" Then
txtTotal2.text = val(txtDiff.text) - val(txtTotal1.text)
End If
End Sub

Private Sub txttotal2_Change()
If Me.TxtModFlg <> "R" Then
If val(txtTotal2.text) > 0 Then
txtinstranc.text = txtTotal2.text
Else
txtinstranc.text = 0
End If
End If
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

Private Sub txtWater_Change()
txtTotal1.text = val(Txtcommission.text) + val(TxtWater.text) + val(txtinstrunce.text)
End Sub


Private Sub VSFlexGrid2_AfterEdit(ByVal row As Long, ByVal Col As Long)
With VSFlexGrid2
Select Case .ColKey(Col)
Case "payed"
If .Cell(flexcpChecked, row, .ColIndex("payed")) = flexChecked Then
.TextMatrix(row, .ColIndex("TransPayedValue")) = .TextMatrix(row, .ColIndex("RemainingValue"))
Else
.TextMatrix(row, .ColIndex("TransPayedValue")) = 0
End If
End Select
End With
RelineOwner2
End Sub
Private Sub VSFlexGrid2_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid2
Select Case .ColKey(Col)
Case "NoteDate"
Cancel = True
Case "ContNoteSerial1"
Cancel = True
Case "branch_name"
Cancel = True
Case "CusName"
Cancel = True
Case "UntName"
Cancel = True
Case "aqarname"
Cancel = True
Case "PayedValue"
Cancel = True
Case "RemainingValue"
Cancel = True
Case "value"
Cancel = True
End Select
End With
End Sub

Private Sub VSFlexGrid3_AfterEdit(ByVal row As Long, ByVal Col As Long)
With VSFlexGrid3
Select Case .ColKey(Col)
Case "payed"
If .Cell(flexcpChecked, row, .ColIndex("payed")) = flexChecked Then
.TextMatrix(row, .ColIndex("TransPayedValue")) = .TextMatrix(row, .ColIndex("RemainingValue"))
Else
.TextMatrix(row, .ColIndex("TransPayedValue")) = 0
End If
End Select
End With
RelineOwner2
End Sub

Private Sub VSFlexGrid3_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid3
Select Case .ColKey(Col)
Case "NoteDate"
Cancel = True
Case "ContNoteSerial1"
Cancel = True
Case "branch_name"
Cancel = True
Case "CusName"
Cancel = True
Case "UntName"
Cancel = True
Case "aqarname"
Cancel = True
Case "PayedValue"
Cancel = True
Case "RemainingValue"
Cancel = True
Case "value"
Cancel = True

End Select
End With
End Sub

Private Sub XPBtnMove_Click(index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case index

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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim RsDev As ADODB.Recordset
    Dim i As Integer

    On Error GoTo ErrTrap
    Option4.value = False
    Option5.value = False
    Option6.value = False
    Option7.value = False
    Fra(2).Visible = False

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

    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    End If

    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
 
    Me.TXT_order_no.text = IIf(IsNull(rs("Order_no").value), "", rs("Order_no").value)
    Me.Text1.text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    XPTxtID1.text = IIf(IsNull(rs("AdvanceID").value), "", (rs("AdvanceID").value))
    txtTransferExpenses.text = IIf(IsNull(rs("TransferExpenses").value), "", (rs("TransferExpenses").value))
         Me.TxtFilterNo.text = IIf(IsNull(rs("FilterID").value), "", rs("FilterID").value)
       Me.TXtFilter.text = IIf(IsNull(rs("FIlterTotal").value), "", rs("FIlterTotal").value)
       
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
 DcboCreditSide.BoundText = IIf(IsNull(rs("CreditSide").value), "", rs("CreditSide").value)
    DcboDebitSide.BoundText = IIf(IsNull(rs("DebitSide").value), "", rs("DebitSide").value)
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(45).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    txt_general_des.text = IIf(IsNull(rs("general_des_notes").value), "", rs("general_des_notes").value)

    txtperson.text = IIf(IsNull(rs("person").value), "", rs("person").value)

    XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", (rs("Note_Value").value))
    dcproject.BoundText = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
    XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
''//
  FrmPriodDate.value = IIf(IsNull(rs("FrmPriodDate").value), Date, rs("FrmPriodDate").value)
    FrmPriodDateH.value = IIf(IsNull(rs("FrmPriodDateH").value), ToHijriDate(FrmPriodDate.value), rs("NoteDateH").value)
      XPDtbTrans.value = IIf(IsNull(rs("ToPriodDate").value), Date, rs("ToPriodDate").value)
    ToPriodDateH.value = IIf(IsNull(rs("ToPriodDateH").value), ToHijriDate(ToPriodDate.value), rs("NoteDateH").value)
    
'''/
    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    Txt_DateHigri.value = IIf(IsNull(rs("NoteDateH").value), ToHijriDate(XPDtbTrans.value), rs("NoteDateH").value)
txtTotalinsuranceS.text = IIf(IsNull(rs("TotalInsurances").value), "", Trim(rs("TotalInsurances").value))
    DCboCashType.ListIndex = IIf(IsNull(rs("CashingType").value), -1, rs("CashingType").value)

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
 ''//
 
 
    optAdd.value = True
   TxtOfficeValue.text = IIf(IsNull(rs("OfficeValue").value), 0, rs("OfficeValue").value)
   TxtOfficeValueNet.text = IIf(IsNull(rs("OfficeValueNet").value), 0, rs("OfficeValueNet").value)
   TxtOfficeValueDiscAdd.text = IIf(IsNull(rs("OfficeValueDiscAdd").value), 0, rs("OfficeValueDiscAdd").value)
   
    If val(rs!AddValue & "") = 1 Then
        optAdd.value = True
    ElseIf val(rs!AddValue & "") = 2 Then
        optDisc.value = True
    End If
    
   TxtNetPayments.text = IIf(IsNull(rs("RenterValue").value), 0, rs("RenterValue").value)
   TxtValuExpenses.text = IIf(IsNull(rs("ExpValue").value), 0, rs("ExpValue").value)
   
   TxtTotalPayedOpBalance.text = IIf(IsNull(rs("TotalPayedOpBalance").value), 0, rs("TotalPayedOpBalance").value)
   TotalPayments.text = IIf(IsNull(rs("TotalPayments").value), 0, rs("TotalPayments").value)
   txtPercent.text = IIf(IsNull(rs("Percentage").value), 0, rs("Percentage").value)
   TxtNetValue.text = IIf(IsNull(rs("NetValue").value), 0, rs("NetValue").value)
   TxtNetPayments.text = IIf(IsNull(rs("RenterValue").value), 0, rs("RenterValue").value)
   TxtPreBalaValue.text = IIf(IsNull(rs("PreBalaValue").value), 0, rs("PreBalaValue").value)
   TxtPreBalaPayed.text = IIf(IsNull(rs("PreBalaPayed").value), 0, rs("PreBalaPayed").value)
   TxtPreBalaRemain.text = IIf(IsNull(rs("PreBalaRemain").value), 0, rs("PreBalaRemain").value)
   TxtPreBalaTransPyed.text = IIf(IsNull(rs("PreBalaTransPyed").value), 0, rs("PreBalaTransPyed").value)
   TxtPreBalaNet.text = IIf(IsNull(rs("PreBalaNet").value), 0, rs("PreBalaNet").value)
   
         Me.Txtcommission.text = IIf(IsNull(rs("commission").value), "", rs("commission").value)
         Me.TxtCommissionOut.text = IIf(IsNull(rs("CommissionOut").value), "", rs("CommissionOut").value)
         Me.TxtRent.text = IIf(IsNull(rs("rent").value), "", rs("rent").value)
         Me.TxtWater.text = IIf(IsNull(rs("Water").value), "", rs("Water").value)
         Me.txtinstrunce.text = IIf(IsNull(rs("Instrunce").value), "", rs("Instrunce").value)
         Me.txtComisin.text = IIf(IsNull(rs("comX").value), "", rs("comX").value)
         Me.txtinstranc.text = IIf(IsNull(rs("ComY").value), "", rs("ComY").value)
         Me.txtComisinold.text = IIf(IsNull(rs("comXold").value), "", rs("comXold").value)
         Me.txtinstrancold.text = IIf(IsNull(rs("ComYold").value), "", rs("ComYold").value)
           Me.TxtNotSreail1.text = IIf(IsNull(rs("NoteOrBonSereal").value), "", rs("NoteOrBonSereal").value)
         Me.TxtNotID.text = IIf(IsNull(rs("NoteOrBonID").value), "", rs("NoteOrBonID").value)
         Me.TxtNotVal.text = IIf(IsNull(rs("NoteOrBonValue").value), "", rs("NoteOrBonValue").value)
 ''/
        
        Me.DcbIqara.BoundText = IIf(IsNull(rs("akarid").value), "", rs("akarid").value)
        Me.DcbIqara2.BoundText = IIf(IsNull(rs("IqarID2").value), IIf(IsNull(rs("akarid").value), "", rs("akarid").value), rs("IqarID2").value)
        Me.DcbUnitType.BoundText = IIf(IsNull(rs("UnitType").value), "", rs("UnitType").value)
        Me.DcbUnitNo.BoundText = IIf(IsNull(rs("UnitNo").value), "", rs("UnitNo").value)
        If Not IsNull(rs("ComResid").value) Then
        If (rs("ComResid").value) = 1 Then
        ComResid(1).value = True
        Else
        ComResid(0).value = True
        End If
        Else
        ComResid(0).value = True
        End If
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
    
    End If

    If DCboCashType.ListIndex = 3 Then
        DBCboClientName.BoundText = IIf(IsNull(rs("projectAccountCode").value), 0, rs("projectAccountCode").value)

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

    ElseIf DCboCashType.ListIndex = 5 Or DCboCashType.ListIndex = 7 Then
        DBCboClientName.BoundText = IIf(IsNull(rs("BTCashAccountcode").value), 0, rs("BTCashAccountcode").value)
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
    End If
XPTxtVal.text = IIf(IsNull(rs("Note_Value2").value), IIf(IsNull(rs("Note_Value").value), 0, (rs("Note_Value").value)) - IIf(IsNull(rs("OfficeValue").value), 0, (rs("OfficeValue").value)), (rs("Note_Value2").value))
    '---------------------------------------------------------------------------

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
RetriveOwnerPaymentData
RetriveOwnerPaymentData202
RetriveOwnerPaymentData203
            If DCboCashType.ListIndex = 8 Then
                ' Frame12(1).Visible = True
            Else
                    Frame12(1).Visible = False
            End If
    '-----------------------------------------------------------------------------
    If DcboCreditSide.text = "" And DcboDebitSide.text = "" Then
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
    End If
Frame5.Visible = False
optAdd_Click
    '-----------------------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim StrTemp As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

  '   On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

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
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·„œ›Ê⁄«  "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboCashType.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If

        If DBCboClientName.text = "" Then
           ' Msg = "ÌÃ» «Œ Ì«—«·«”„"
          '  MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            DBCboClientName.SetFocus
           ' SendKeys "{F4}"
           ' Exit Sub
        End If

        If XPTxtVal.text = "" Then
            Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·„œ›Ê⁄«  "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtVal.SetFocus
            Exit Sub
        End If

        If Not IsNumeric(XPTxtVal.text) Then
            Msg = "ﬁÌ„… «·„œ›Ê⁄«  ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtVal.SetFocus
            SelectText XPTxtVal
            Exit Sub
        End If

        If Me.CboPayMentType.ListIndex = -1 Then
            Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPayMentType.SetFocus
            Exit Sub
        End If
 If Me.CboPayMentType.ListIndex = 6 Then
 If TxtFilterNo.text = "" Then
            Msg = "ÌÃ» ≈Œ Ì«—  —ﬁ„ «· ’›ÌÂ ...!!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtFilterNo.SetFocus
           Exit Sub
             End If
        End If
         If Me.CboPayMentType.ListIndex = 7 Then
 If TxtNotSreail1.text = "" Then
            Msg = "ÌÃ» ≈Œ Ì«—  —ﬁ„ ”‰œ ﬁ»÷ «·⁄—»Ê‰ ...!!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtNotSreail1.SetFocus
           Exit Sub
             End If
        End If
        If Me.CboPayMentType.ListIndex = 0 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBox.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPayMentType.ListIndex = 1 Or Me.CboPayMentType.ListIndex = 3 Then

            If Me.DcboBankName.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
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
        
        ElseIf Me.CboPayMentType.ListIndex = 2 Then

            If Me.DcboBankName.BoundText = "" Then
                Msg = "Õœœ   «·»‰ﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
                Msg = "Õœœ —ﬁ„ «·ÕÊ«·Â...!!"
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
                Msg = "»—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·›« Ê—…..!!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                CboTrans.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim(Me.TxtTransSerial.text) = "" Then
                Msg = "»—Ã«¡ ≈œŒ«· —ﬁ„ «·›« Ê—…..!!!"
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
            If Voucher_coding(val(my_branch), XPDtbTrans.value, 4, 5) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ œ›⁄ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbTrans.value, 4, 5) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 4, 5)
                End If
            End If
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
            Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
        ElseIf TxtModFlg.text = "E" Then
         StrSQL = "Delete  TblAqarCommissions  where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If

        '         rs("AdvanceID").value = Val(XPTxtID1.text)
          If DCboCashType.ListIndex = 6 Then
            rs("FilterID").value = IIf(TxtFilterNo.text = "", Null, TxtFilterNo.text)
           rs("FIlterTotal").value = IIf(TXtFilter.text = "", Null, TXtFilter.text)
            Else
            rs("FilterID").value = Null
           rs("FIlterTotal").value = Null
        End If
               If DCboCashType.ListIndex = 7 Then
               GetNotesSalesInformation val(TxtNotID.text)
                   rs("commission").value = IIf(Txtcommission.text = "", 0, Trim(Txtcommission.text))
              rs("CommissionOut").value = IIf(Me.TxtCommissionOut.text = "", 0, Trim(TxtCommissionOut.text))
        rs("rent").value = IIf(TxtRent.text = "", 0, Trim(TxtRent.text))
        rs("Water").value = IIf(TxtWater.text = "", 0, Trim(TxtWater.text))
        rs("Instrunce").value = IIf(txtinstrunce.text = "", 0, Trim(txtinstrunce.text))
        rs("comX").value = IIf(txtComisin.text = "", 0, Trim(txtComisin.text))
        rs("ComY").value = IIf(txtinstranc.text = "", 0, Trim(txtinstranc.text))
        rs("comXold").value = IIf(txtComisinold.text = "", 0, Trim(txtComisinold.text))
        rs("ComYold").value = IIf(txtinstrancold.text = "", 0, Trim(txtinstrancold.text))
            rs("NoteOrBonSereal").value = IIf(TxtNotSreail1.text = "", Null, TxtNotSreail1.text)
           rs("NoteOrBonID").value = IIf(TxtNotID.text = "", Null, TxtNotID.text)
           rs("NoteOrBonValue").value = IIf(TxtNotVal.text = "", Null, TxtNotVal.text)
            Else
            rs("NoteOrBonSereal").value = Null
           rs("NoteOrBonID").value = Null
            rs("NoteOrBonValue").value = Null
        End If
         
        rs("Order_no").value = IIf(Trim(Me.TXT_order_no.text) = "", Null, Trim(Me.TXT_order_no.text))
        rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
        rs("RenterValue").value = val(TxtNetPayments.text)
        rs("ExpValue").value = val(TxtValuExpenses.text)
        rs("OfficeValue").value = val(TxtOfficeValue.text)
       rs("OfficeValueDiscAdd").value = val(TxtOfficeValueDiscAdd.text)
       rs("OfficeValueNet").value = val(TxtOfficeValueNet.text)
        
        rs("TotalPayments").value = val(TotalPayments.text)
        rs("TotalPayedOpBalance").value = val(TxtTotalPayedOpBalance.text)
        rs("Percentage").value = val(txtPercent.text)
        rs("NetValue").value = val(TxtNetValue.text)
        
        rs("PreBalaValue").value = val(TxtPreBalaValue.text)
        rs("PreBalaPayed").value = val(TxtPreBalaPayed.text)
        rs("PreBalaRemain").value = val(TxtPreBalaRemain.text)
        rs("PreBalaTransPyed").value = val(TxtPreBalaTransPyed.text)
        rs("PreBalaNet").value = val(TxtPreBalaNet.text)
        
        If ComResid(1).value = True Then
        rs("ComResid").value = 1
        Else
        rs("ComResid").value = 0
        End If
        
        If optAdd.value = True Then
            rs("AddValue").value = 1
        ElseIf optDisc.value = True Then
            rs("AddValue").value = 2
        End If
   
    


        rs("branch_no").value = val(Me.dcBranch.BoundText)
        rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, val(XPTxtVal.text))
        rs("note_value_by_characters").value = IIf(lbl(18).Caption = "", Null, lbl(18).Caption)
     
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))
        rs("general_des_notes").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    
        rs("person").value = IIf(Me.txtperson.text = "", "", Me.txtperson.text)

        rs("NoteType").value = 5
        rs("TransferExpenses").value = val(txtTransferExpenses.text)
        'TransferExpenses
    ''//
     rs("CreditSide").value = IIf(Trim(DcboCreditSide.BoundText) = "", Null, (DcboCreditSide.BoundText))
     rs("DebitSide").value = IIf(Trim(DcboDebitSide.BoundText) = "", Null, (DcboDebitSide.BoundText))
    rs("FrmPriodDate").value = Me.FrmPriodDate.value
    rs("FrmPriodDateH").value = Me.FrmPriodDateH.value
    rs("ToPriodDate").value = Me.ToPriodDate.value
    rs("ToPriodDateH").value = Me.ToPriodDateH.value
    rs("IqarID2").value = val(Me.DcbIqara2.BoundText)
    rs("TotalInsurances").value = IIf(txtTotalinsuranceS.text = "", Null, txtTotalinsuranceS.text)
    '''/
          rs("NoteDate").value = XPDtbTrans.value
       ' rs("NoteDate").value = Format$(Date, "dd-mm-yyyy")
        rs("NoteDateH").value = Me.Txt_DateHigri.value
   
        rs("CashingType").value = IIf(DCboCashType.ListIndex = -1, Null, DCboCashType.ListIndex)

        If DCboCashType.ListIndex = 3 Then
            rs("projectAccountCode").value = IIf(DBCboClientName.text = "", Null, DBCboClientName.BoundText)

        ElseIf DCboCashType.ListIndex = 4 Then
            rs("EmpAccountCode").value = IIf(DBCboClientName.text = "", Null, DBCboClientName.BoundText)
            rs("CusID").value = Null

            rs("person").value = IIf(DBCboClientName.text = "", "", Trim(DBCboClientName.text))
        
        ElseIf DCboCashType.ListIndex = 5 Then
            rs("BTCashAccountcode").value = IIf(DBCboClientName.text = "", Null, DBCboClientName.BoundText)
            rs("CusID").value = Null
        Else
            rs("CusID").value = IIf(DBCboClientName.text = "", Null, DBCboClientName.BoundText)
 
        End If
    
        If Option4.value = True Then
            rs("salary_or_advance").value = 0
        ElseIf Option5.value = True Then
            rs("salary_or_advance").value = 1

            If val(XPTxtID1.text) = 0 Then
                XPTxtID1.text = CStr(new_id("TblEmpAdvance", "AdvanceID", "", True))
            End If

            rs("AdvanceID").value = val(XPTxtID1.text)
    
        ElseIf Option6.value = True Then
            rs("salary_or_advance").value = 2
        ElseIf Option7.value = True Then
            rs("salary_or_advance").value = 3
              
        Else
            rs("salary_or_advance").value = Null
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

        End If
        rs("akarid").value = IIf(val(Me.DcbIqara.BoundText) <> 0, val(DcbIqara.BoundText), Null)
        rs.Fields("UnitType").value = IIf(Me.DcbUnitType.BoundText <> "", val(DcbUnitType.BoundText), Null)
        rs.Fields("UnitNo").value = IIf(Me.DcbUnitNo.BoundText <> "", val(DcbUnitNo.BoundText), Null)
        
        rs("UserID").value = user_id
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        rs("foxy_no").value = val(Text1.text)
        rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        rs("numbering_type1").value = sand_numbering_type(4) '”‰œ «·œ›⁄
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)
        rs("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
    
        rs.update
 Dim IarType As Integer
            IarType = AqarCommisionType(val(DcbIqara.BoundText))
If SystemOptions.NoCreatJLInRentContract = True And DCboCashType.ListIndex = 6 Then
If IarType <> 0 Then
OtherOwnerNoreatJlInContractFiter 1, val(XPTxtID.text)
Else
MyOwnerNoreatJlInContractFiter 1, val(XPTxtID.text)
End If
GoTo llx:
End If

        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            Set RsDev = New ADODB.Recordset
        '    RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                            StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
            Line1 = setfoxy_Line
            Line2 = setfoxy_Line
            Line3 = setfoxy_Line

            '«·ÿ—› «·„œÌ‰
            ' ›Ì Õ«·… «·ÕÊ«·«  «·»‰ﬂÌ… ÊÊÃÊœ „’—Ê›«  »‰ﬂÌ… ⁄·Ì⁄«
            If CboPayMentType.ListIndex = 2 And val(Me.txtTransferExpenses.text) > 0 Then
                RsDev.AddNew
                RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                RsDev("Double_Entry_Vouchers_ID").value = LngDevID
                RsDev("DEV_ID_Line_No").value = 2
                RsDev("DEV_ID_Line_No1").value = Line2
                RsDev("Account_Code").value = Account_Code_dynamic
                RsDev("Value").value = val(Me.txtTransferExpenses.text)
                RsDev("Credit_Or_Debit").value = 0
                RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & txt_general_des
                RsDev("Notes_ID").value = val(XPTxtID.text)
                RsDev("RecordDate").value = Me.XPDtbTrans.value
                RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
                RsDev("UserID").value = Me.DCboUserName.BoundText
                RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            
                RsDev.update
        
            End If

            '44444444444444444444444444
       ' Set RsDev = New ADODB.Recordset
       Dim LastLine As Integer
      If DCboCashType.ListIndex = 8 Then
       LastLine = payGlPaymentOwner(LngDevID, val(XPTxtID.text))
     GoTo llx
     End If
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 1
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.text)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & txt_general_des ' XPMTxtRemarks.text
            
            If DCboCashType.ListIndex = 3 Then
                Dim project_id As Integer
                project_id = get_project_id(DBCboClientName.BoundText, "expanses_account")
                RsDev("project_id").value = project_id
                RsDev("Double_Entry_Vouchers_Description").value = "’—› ⁄·Ï „‘—Ê⁄" & DBCboClientName.text
            End If
            
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            
            RsDev.update
           
            '«·ÿ—› «·œ«∆‰
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.text) + val(Me.txtTransferExpenses.text)
            RsDev("Credit_Or_Debit").value = 1
            
            If CboPayMentType.ListIndex = 2 And val(Me.txtTransferExpenses.text) > 0 Then
                RsDev("DEV_ID_Line_No").value = 3
                RsDev("DEV_ID_Line_No1").value = Line3
            Else
                RsDev("DEV_ID_Line_No").value = 2
                RsDev("DEV_ID_Line_No1").value = Line2
            End If

            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.text & CHR(13) & txt_general_des ' XPMTxtRemarks.text
            ' RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            
            RsDev("Notes_ID").value = val(XPTxtID.text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
        
            LblDevID.Caption = LngDevID
            lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
        End If
llx:
        saveChequeBoxContents1 (val(XPTxtID.text))
        SaveOwnerPayment
        SaveOwnerPayment202
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        CuurentLogdata

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

   
         If val(DCboCashType.ListIndex) = 8 Then
           updateNotesValueAndNobytext val(XPTxtID.text), Format(XPTxtVal.text, "###.00")
           Else
           updateNotesValueAndNobytext val(XPTxtID.text), Format(XPTxtVal.text, "###.00")
       End If
            
            
        '«· Ê“Ì⁄ ⁄·Ï „—ﬂ“ «· ﬂ·›… «·⁄«„
           If Me.DcCostCenter.BoundText <> "" Then
        save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.text, "„œ›Ê⁄« ", Me.XPDtbTrans.value
        save_cost_center
  
            End If
        
        'Õ›Ÿ «·„’«—Ì› › ÃœÊ· «·„œ›Ê⁄«  Ê «·„ﬁ»Ê÷« 
     
     '   If SavePaymentAndReciveDetails(0, TxtNoteSerial.text, txtNoteSerial1.text, txt_ORDER_NO.text, XPDtbTrans.value) = True Then
       ' End If

        'Õ›Ÿ »Ì«‰«  «·”·›…
        saveAdvancedData
        
    End If

    WriteCustomerBalPublic Me.DcboDebitSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString

    WriteInfo

    If Option1.value = True Then
        FIFO_FUNCTION val(DBCboClientName.BoundText)
    End If
   
    If Option2.value And lblsqlstring <> "Label1" And lblsqlstring <> "" Then
        Distribute_to_bills Me.lblsqlstring, val(DBCboClientName.BoundText)
    End If
    rs.Resync adAffectCurrent
     TxtModFlg.text = "R"
     
   Retrive val(XPTxtID.text)
  
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
Function payGlPaymentOwner(LngDevID As Long, notes_id As Double) As Double

  
 If DCboCashType.ListIndex <> 8 Then Exit Function
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
     BranchID = val(Me.dcBranch.BoundText)
Dim i As Integer
Line1 = 3
total_value = val(TxtNetPayments.text)
If total_value <= 0 Then
    With VSFlexGrid1

        For i = .FixedRows To .rows - 1
 
            If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked And val(.TextMatrix(i, .ColIndex("TransPayedValue"))) > 0 Then
             BranchID = val(Me.dcBranch.BoundText)
             BranchID2 = val(.TextMatrix(i, .ColIndex("BranchId")))
             DeptSide = getBranchCurrentAccount(BranchID)
             credit_side = getBranchCurrentAccount(BranchID2)
              DeptSide1 = DcboDebitSide.BoundText
              CreditSide1 = DcboCreditSide.BoundText
                                                 
                total_value = Round(.TextMatrix(i, .ColIndex("TransPayedValue")), 2)
                CURRENT_LINE = setfoxy_Line

                If total_value > 0 Then
                
              Msg = "  ”œ«œ Ã“¡ „‰ œ›⁄… —ﬁ„" & CHR(13) & .TextMatrix(i, .ColIndex("PaymentNo"))
              Msg = Msg & CHR(13) & "··⁄ﬁ«— " & .TextMatrix(i, .ColIndex("aqarname"))
              Msg = Msg & CHR(13) & "··›—⁄ " & .TextMatrix(i, .ColIndex("branch_name"))
              Msg = Msg & CHR(13) & "”œœ  „‰  " & dcBranch.text
                          
                                      '„«·ﬂ
                                        If ModAccounts.AddNewDev(LngDevID, Line1, DeptSide1, total_value, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID2, , , , , , , , , , val(.TextMatrix(i, .ColIndex("Aqarid")))) = False Then
                                                                   
                                                              End If
                                                              
                                                              Line1 = Line1 + 1
                                                  '⁄Âœ…
                                                  
                                                  If ModAccounts.AddNewDev(LngDevID, Line1, CreditSide1, total_value, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , val(.TextMatrix(i, .ColIndex("Aqarid")))) = False Then
                                                              End If
                                                              
                                                                      Line1 = Line1 + 1
               
                                
            End If
                             
                     
End If
        Next i
    End With
Else

              DeptSide1 = DcboDebitSide.BoundText
              CreditSide1 = DcboCreditSide.BoundText
                                  
                
                CURRENT_LINE = setfoxy_Line
total_value = val(TxtOfficeValueNet.text)

                If total_value > 0 And SystemOptions.NoCreatJLInRentContract = True Then
                Dim RentAccount As String
                Dim dummyCommissionAcc As String
                                      '„«·ﬂ
                                      Msg = "”‰œ ’—› „œ›Ê⁄«  ··„«·ﬂ  «À»«  ⁄„Ê·Â «·„ﬂ »"
                                      Msg = Msg & " ··⁄ﬁ«— " & DcbIqara2.text
                                     Msg = Msg & " —ﬁ„ «·”‰œ " & Me.TxtNoteSerial1.text
                                     
                                      total_value = total_value ' / 1.05
                                        If ModAccounts.AddNewDev(LngDevID, Line1, DeptSide1, total_value, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , val(Me.DcbIqara2.BoundText)) = False Then
                                                                   
                                                              End If
                                                              
                                                              Line1 = Line1 + 1
                                                              total_value = total_value / 1.05
                                             dummyCommissionAcc = get_account_code_branch(207, Me.dcBranch.BoundText)
                                          If ModAccounts.AddNewDev(LngDevID, Line1, dummyCommissionAcc, total_value, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , val(Me.DcbIqara2.BoundText)) = False Then
                                                   End If
                                                              
                                             Line1 = Line1 + 1
                                             
                                                  '⁄Âœ…
                                                  
                                                  total_value = total_value * 5 / 100
                                             '     If ModAccounts.AddNewDev(LngDevID, Line1, DeptSide1, total_value, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , val(Me.DcbIqara2.BoundText)) = False Then
                                             '      End If
                                             '     Line1 = Line1 + 1
                                             Msg = "«·ﬁÌ„Â «·„÷«›… ·⁄„Ê·«   «·„ﬂ » "
                                     
                                      Msg = Msg & " ··⁄ﬁ«— " & DcbIqara2.text
                                     Msg = Msg & " —ﬁ„ «·”‰œ " & Me.TxtNoteSerial1.text
                                     
                                                  GetValueAddedAccount XPDtbTrans.value, , RentAccount, 1, 21
                                             
                                                      If ModAccounts.AddNewDev(LngDevID, Line1, RentAccount, total_value, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , val(Me.DcbIqara2.BoundText)) = False Then
                                                   End If
                                                              
                                             Line1 = Line1 + 1
               
                                
             End If
         ''SALIMHERE         total_value = val(TxtNetPayments.Text)
                     total_value = val(TxtNetValue.text)
             
                If total_value > 0 Then
                Msg = "”‰œ ’—› „” Õﬁ«  «·„«·ﬂ "
                        
                                      Msg = Msg & " ··⁄ﬁ«— " & DcbIqara2.text
                                     Msg = Msg & " —ﬁ„ «·”‰œ " & Me.TxtNoteSerial1.text
                                     
                                      '„«·ﬂ
                                        If ModAccounts.AddNewDev(LngDevID, Line1, DeptSide1, total_value, 0, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , val(Me.DcbIqara2.BoundText)) = False Then
                                                                   
                                                              End If
                                                              
                                                              Line1 = Line1 + 1
                                                  '⁄Âœ…
                                                  
                                                  If ModAccounts.AddNewDev(LngDevID, Line1, CreditSide1, total_value, 1, Msg, val(notes_id), , , , XPDtbTrans.value, user_id, , , , , , , , , CURRENT_LINE, , , , , , , , , BranchID, , , , , , , , , , val(Me.DcbIqara2.BoundText)) = False Then
                                                   End If
                                                              
                                             Line1 = Line1 + 1
               
                                
             End If
                             
   End If

             
 payGlPaymentOwner = Line1 + 1
ErrTrap:
End Function
Function saveAdvancedData()

    Dim StrSQL  As String
    StrSQL = "Delete From TblEmpAdvance Where AdvanceID=" & val(Me.XPTxtID1.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
         
    StrSQL = "Delete From TblEmpAdvanceDetails Where AdvanceID=" & val(Me.XPTxtID1.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
     
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
    
    
    For i = Me.Fg.FixedRows To Fg.rows - 1

        If val(Fg.TextMatrix(i, Fg.ColIndex("PartNO"))) <> 0 Then
            RsDetails.AddNew
            RsDetails("AdvanceID").value = val(XPTxtID1.text)
            RsDetails("PartNO").value = Fg.TextMatrix(i, Fg.ColIndex("PartNO"))
            RsDetails("PartValue").value = Fg.TextMatrix(i, Fg.ColIndex("PartValue"))
            RsDetails("PartDate").value = Fg.TextMatrix(i, Fg.ColIndex("PartDate"))
            RsDetails.update
        End If

    Next i

End Function

Function saveChequeBoxContents1(NoteID As Double)

    If SystemOptions.banks_Accounts3 = False Then Exit Function
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
        rs("ChequeValue").value = val(XPTxtVal.text)
    
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
  
        Add_new_notes Me.XPDtbTrans, 2001, current_value, Rs3("transactionsid").value, CusID, DcboBox.BoundText, 1, val(DCboUserName.BoundText)
  
        Rs3.MoveNext
    Next i

    ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
    txtAdv_payment_value.text = total_value
    change_adv_payment_value XPTxtID.text, total_value
    ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
    Rs3.Close

End Function

Function Add_new_notes(NoteDate As Date, NoteType As Integer, Note_Value As Double, Transaction_ID As Integer, CusID As Double, BoxID As Integer, displayed As Integer, UserID As Integer)
    Dim RsDev As New ADODB.Recordset
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
    
    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then

                CuurentLogdata ("D")

                rs.delete
                Dim StrSQL As String
             '   StrSQL = "Delete From notes  Where (NoteType=2001 OR NoteType=5 ) AND NoteSerial=" & val(TxtNoteSerial.Text)
             '   Cn.Execute StrSQL, , adExecuteNoRecords
        
                StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
       
                StrSQL = "Delete From ReciveDetails Where NoteSerial1='" & val(TxtNoteSerial1.text) & "'"
                Cn.Execute StrSQL, , adExecuteNoRecords
    
                StrSQL = "Delete  TblAqarCommissions  where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
    
                StrSQL = "Delete From TblEmpAdvance Where AdvanceID=" & val(Me.XPTxtID1.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
         
                StrSQL = "Delete From TblEmpAdvanceDetails Where AdvanceID=" & val(Me.XPTxtID1.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = " Delete From TblNotesOwnerPayment Where NoteID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblOwnerPayment Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                StrSQL = " Delete From TblNotesOwnerPayment202 Where NoteID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblOwnerPayment202 Where   NoteID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
              DeleteBillBuy
              DeleteOwner202
              DeleteOwner203
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
'Private Sub DcbUnitNo_Change()
'If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
'Dim str As String
'If DCboCashType.ListIndex <> 7 Then Exit Sub
' str = checkDepositeRent(val(DcbUnitNo.BoundText), XPDtbTrans)
'
'
'If str <> "" Then
'MsgBox str, vbInformation
'End If
'
'End If
'End Sub

'Private Sub DcbUnitType_Change()
'Dim Dcombos As ClsDataCombos
'Dim idd As Long
'Dim idd1 As Long
'   Set Dcombos = New ClsDataCombos
'  ' Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
'
'If val(DcbIqara.BoundText) > 0 Then
'idd = val(DcbIqara.BoundText)
'
'idd1 = val(DcbUnitType.BoundText)
'If Me.TxtModFlg = "R" Then
'Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
'Else
'Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo
'End If
'End If
'End Sub

'Private Sub DcbUnitType_Click(Area As Integer)
'DcbUnitType_Change
'End Sub
Sub RelineOwner2()
    Dim IntCounter As Integer
    Dim Sm As Double
    Dim Sm1 As Double
    Dim suminurancse As Double
    Sm = 0
    suminurancse = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid2
        For i = .FixedRows To .rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
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
           suminurancse = suminurancse + val(.TextMatrix(i, .ColIndex("insuranceValue")))
           End If
           Next i
  
    End With
    
        Sm1 = 0
    IntCounter = 0
   
    With Me.VSFlexGrid3
        For i = .FixedRows To .rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
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
           Sm1 = Sm1 + val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           End If
           Next i
  
    End With
    Dim total As Double
    Dim AmolValue As Double
    AmolValue = GetAmolOwnerIaqar(val(DBCboClientName.BoundText)) / 100
    TxtTotalPayedOpBalance.text = val(TxtPreBalaTransPyed.text)
    TotalPayments.text = Sm
    txtPercent.text = AmolValue
    total = val(TotalPayments.text) + val(TxtTotalPayedOpBalance.text)

    TxtOfficeValue.text = (total - suminurancse) * 1.05 * val(txtPercent.text)
    
    TxtOfficeValue.text = (total - suminurancse) * 1.05 * val(txtPercent.text)
    
If optAdd Then
    TxtOfficeValueNet = val(TxtOfficeValueDiscAdd) + val(TxtOfficeValue)
Else
    TxtOfficeValueNet = val(TxtOfficeValue) - val(TxtOfficeValueDiscAdd)
End If
TxtNetPayments.text = val(TotalPayments.text) + val(TxtTotalPayedOpBalance.text) - val(TxtOfficeValueNet.text)
' TxtNetPayments.Text = Total - Total * 1.05 * val(TxtPercent.Text)
TxtNetPayments.text = total - val(TxtOfficeValueNet.text)


    TxtValuExpenses.text = Sm1
    TxtNetValue.text = val(TxtNetPayments) - val(TxtValuExpenses)
    XPTxtVal.text = TxtNetValue.text
   XPTxtVal.Enabled = False
End Sub
Function GetAmolOwnerIaqar(Optional CusID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     AmolaValus"
sql = sql & " From dbo.TblAqar"
'sql = sql & " Where (ownerid = " & CusID & ") And (Not (AmolaValus Is Null))"
sql = sql & " Where (Aqarid = " & val(DcbIqara2.BoundText) & ") And (Not (AmolaValus Is Null))"


rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetAmolOwnerIaqar = IIf(IsNull(rs2("AmolaValus").value), 1, rs2("AmolaValus").value)
Else
GetAmolOwnerIaqar = 1
End If
End Function
Sub RelineBu22()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid1
        For i = .FixedRows To .rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
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
End Sub
Private Sub VSFlexGrid1_AfterEdit(ByVal row As Long, ByVal Col As Long)
With VSFlexGrid1
Select Case .ColKey(Col)
Case "payed"
If .Cell(flexcpChecked, row, .ColIndex("payed")) = flexChecked Then
.TextMatrix(row, .ColIndex("TransPayedValue")) = .TextMatrix(row, .ColIndex("RemainingValue"))
Else
.TextMatrix(row, .ColIndex("TransPayedValue")) = 0
End If
End Select
End With

RelineBuy
RelineBu22
End Sub
Private Sub VSFlexGrid1_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid1
Select Case .ColKey(Col)
Case "TransPayedValue"
If .Cell(flexcpChecked, row, .ColIndex("payed")) = flexChecked Then
Cancel = False
Else
End If

Case "aqarNo"
Cancel = True
Case "aqarname"
Cancel = True
Case "PaymentNo"
Cancel = True
Case "RecDate"
Cancel = True
Case "AllowDateH"
Cancel = True
Case "branch_name"
Cancel = True
Case "PayedValue"
Cancel = True
Case "RemainingValue"
Cancel = True
Case "value"
Cancel = True

End Select
End With
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
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Sub GetWonerID(Optional Aqarid As Integer = 0)
If Aqarid <> 0 Then
Dim Rs9  As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
Dim sql As String
sql = "select * from tblaqar where Aqarid =" & Aqarid & ""
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs9.RecordCount > 0 Then
Txtownerid.text = IIf(IsNull(Rs9("ownerid").value), 0, Rs9("ownerid").value)
End If
End If
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
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
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
            RetriveAdvanced val(XPTxtID1.text)
        Else
            Fra(2).Visible = False
        End If
    End If

End Sub

Public Sub RetriveAdvanced(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEmpAdvance  Where (TblEmpAdvance.AdvanceType =0) Order By AdvanceID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount < 1 Then
        Fra(2).Visible = False
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Fra(2).Visible = False
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
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.rows = Fg.FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        Fra(2).Visible = True
        RsDetails.MoveFirst
        Fg.rows = Fg.FixedRows + RsDetails.RecordCount

        For i = Me.Fg.FixedRows To Fg.rows - 1
            Fg.TextMatrix(i, Fg.ColIndex("PartNO")) = RsDetails("PartNO").value
            Fg.TextMatrix(i, Fg.ColIndex("PartValue")) = Round(RsDetails("PartValue").value, 2)
            Fg.TextMatrix(i, Fg.ColIndex("PartDate")) = DisplayDate(CDate(RsDetails("PartDate").value))
            RsDetails.MoveNext
        Next i

    End If

    RsDetails.Close
    Set RsDetails = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtVal_Change()
If DCboCashType.ListIndex = 7 Then
txtDiff.text = val(Me.TxtNotVal.text) - val(XPTxtVal.text)
txtTotal2.text = val(txtDiff.text) - val(txtTotal1.text)
If val(txtDiff.text) >= (val(Txtcommission.text) - val(TxtCommissionOut.text)) Then
txtComisin.text = val(Txtcommission.text) - val(TxtCommissionOut.text)
Else
txtComisin.text = val(txtDiff.text)
End If
End If
    'Me.lbl(18).Caption = WriteNo(Me.XPTxtVal.text, 0, True)
    XPTxtValView.text = Format(val(XPTxtVal.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(18).Caption = WriteNo(Format(val(Me.XPTxtVal.text) + val(Me.txtTransferExpenses.text), "0.00"), 0, True, ".", , 0)

    Else
 
        Me.lbl(18).Caption = WriteNo(Format(val(Me.XPTxtVal.text) + val(Me.txtTransferExpenses.text), "0.00"), 0, True, ".", , 1)

    End If

    If TxtModFlg.text = "N" Then
        txtAdv_payment_value.text = XPTxtVal.text
    End If

End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
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

Public Function newrecord()
Cmd_Click (0)
End Function

