VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Frmovers 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄—Ê÷ «·«’‰«›"
   ClientHeight    =   10185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18165
   Icon            =   "FrmOvers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   18165
   Visible         =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   405
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   186
      Top             =   600
      Visible         =   0   'False
      Width           =   3195
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ„Ì«  «Ê«„— «·‘—«¡"
         Height          =   285
         Index           =   1
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   188
         Top             =   0
         Width           =   1125
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—Ê÷ «·«’‰«›"
         Height          =   285
         Index           =   0
         Left            =   1860
         RightToLeft     =   -1  'True
         TabIndex        =   187
         Top             =   120
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.CheckBox opt_Th 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Œ„Ì”"
      Height          =   315
      Left            =   9750
      TabIndex        =   169
      Top             =   960
      Width           =   870
   End
   Begin VB.CheckBox opt_We 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«—»⁄«¡"
      Height          =   315
      Left            =   11190
      TabIndex        =   168
      Top             =   960
      Width           =   840
   End
   Begin VB.CheckBox opt_tu 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·À·«À«¡"
      Height          =   315
      Left            =   12555
      TabIndex        =   167
      Top             =   960
      Width           =   810
   End
   Begin VB.CheckBox opt_mo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«À‰Ì‰"
      Height          =   315
      Left            =   13875
      TabIndex        =   166
      Top             =   960
      Width           =   840
   End
   Begin VB.CheckBox opt_su 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«Õœ"
      Height          =   315
      Left            =   15315
      TabIndex        =   165
      Top             =   960
      Width           =   630
   End
   Begin VB.CheckBox opt_sa 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”» "
      Height          =   315
      Left            =   16245
      TabIndex        =   164
      Top             =   960
      Width           =   810
   End
   Begin VB.CheckBox opt_Fr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Ã„⁄…"
      Height          =   315
      Left            =   8520
      TabIndex        =   163
      Top             =   960
      Width           =   840
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   2820
      Index           =   2
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   6510
      Width           =   17205
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -1440
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   600
         Width           =   855
      End
      Begin VSFlex8Ctl.VSFlexGrid FgItemPloice 
         Height          =   2145
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   16905
         _cx             =   29819
         _cy             =   3784
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
         Rows            =   1
         Cols            =   30
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmOvers.frx":038A
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   12
         Left            =   16320
         TabIndex        =   75
         Top             =   2430
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   476
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
         ButtonImage     =   "FrmOvers.frx":07F3
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   16
         Left            =   240
         TabIndex        =   76
         Top             =   2430
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   476
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·ﬂ·"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmOvers.frx":0D8D
         DrawFocusRectangle=   0   'False
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·«’‰«› «Ê „Ã„Ê⁄«  «·«’‰«› «· Ì Ìÿ»ﬁ ⁄·ÌÂ« «·⁄—÷"
      Height          =   3165
      Index           =   0
      Left            =   18960
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   2400
      Visible         =   0   'False
      Width           =   10125
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -1200
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   600
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DcbGroup 
         Bindings        =   "FrmOvers.frx":1327
         Height          =   315
         Left            =   5280
         TabIndex        =   63
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSDataListLib.DataCombo DcbItem 
         Bindings        =   "FrmOvers.frx":133C
         Height          =   315
         Left            =   5280
         TabIndex        =   64
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
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
      Begin VSFlex8Ctl.VSFlexGrid FgItems 
         Height          =   1395
         Left            =   120
         TabIndex        =   67
         Top             =   1320
         Width           =   9945
         _cx             =   17542
         _cy             =   2461
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
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmOvers.frx":1351
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
      Begin ImpulseButton.ISButton BtonAdd1 
         Height          =   420
         Left            =   4200
         TabIndex        =   71
         Top             =   960
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   741
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "≈œ—«Ã"
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
         ButtonImage     =   "FrmOvers.frx":1425
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   21
         Left            =   9360
         TabIndex        =   74
         Top             =   2760
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   476
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–›"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmOvers.frx":17BF
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   14
         Left            =   120
         TabIndex        =   77
         Top             =   2760
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   476
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·ﬂ·"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmOvers.frx":1D59
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·’‰›"
         Height          =   285
         Index           =   10
         Left            =   8640
         TabIndex        =   66
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„Ã„Ê⁄…"
         Height          =   285
         Index           =   15
         Left            =   8640
         TabIndex        =   65
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   18480
      TabIndex        =   32
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   18720
      TabIndex        =   31
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   14880
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   18420
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   465
      Left            =   -360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   17595
      _cx             =   31036
      _cy             =   820
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
      Caption         =   " ⁄—Ê÷ «·«’‰«›  "
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1425
         TabIndex        =   2
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
         ButtonImage     =   "FrmOvers.frx":22F3
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
         Left            =   360
         TabIndex        =   3
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
         ButtonImage     =   "FrmOvers.frx":268D
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
         Left            =   1950
         TabIndex        =   4
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
         ButtonImage     =   "FrmOvers.frx":2A27
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
         Left            =   885
         TabIndex        =   5
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
         ButtonImage     =   "FrmOvers.frx":2DC1
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   6120
         Picture         =   "FrmOvers.frx":315B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2280
         TabIndex        =   30
         Top             =   480
         Width           =   2205
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   12420
      TabIndex        =   6
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   237240321
      CurrentDate     =   38784
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   4440
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9420
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
         Left            =   7230
         TabIndex        =   8
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   6375
         TabIndex        =   9
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   5535
         TabIndex        =   10
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   4680
         TabIndex        =   11
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   3705
         TabIndex        =   12
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   120
         TabIndex        =   13
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
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
         Left            =   975
         TabIndex        =   14
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
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
         Left            =   2760
         TabIndex        =   25
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
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
         Index           =   9
         Left            =   1920
         TabIndex        =   34
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   " ’œÌ—"
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
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   13440
      TabIndex        =   15
      Top             =   9600
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   18720
      TabIndex        =   16
      Top             =   3570
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   18840
      TabIndex        =   26
      Top             =   1920
      Width           =   1365
      _ExtentX        =   2408
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
   Begin MSDataListLib.DataCombo DcbBranch 
      Bindings        =   "FrmOvers.frx":6DC3
      Height          =   315
      Left            =   7320
      TabIndex        =   28
      Top             =   600
      Width           =   3375
      _ExtentX        =   5953
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   8172
      Left            =   0
      TabIndex        =   35
      Top             =   1320
      Width           =   18120
      _cx             =   31962
      _cy             =   14414
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   0
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "⁄—Ê÷ «·«’‰«›|Õ«·Â «·«⁄ „«œ"
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
      DogEars         =   0   'False
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   1
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Picture(0)      =   "FrmOvers.frx":6DD8
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7710
         Left            =   18765
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   45
         Width           =   18030
         _cx             =   31803
         _cy             =   13600
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
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   3630
            Left            =   120
            TabIndex        =   37
            Tag             =   "1"
            Top             =   240
            Width           =   13230
            _cx             =   23336
            _cy             =   6403
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
            Rows            =   3
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmOvers.frx":7172
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
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   4080
            Width           =   3375
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   9960
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   7710
         Index           =   15
         Left            =   45
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   45
         Width           =   18030
         _cx             =   31803
         _cy             =   13600
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   12
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
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   1
         ChildSpacing    =   1
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
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmOvers.frx":72BE
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7680
            Index           =   16
            Left            =   15
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   15
            Width           =   18000
            _cx             =   31750
            _cy             =   13547
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
            Appearance      =   5
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
            Begin TabDlg.SSTab SSTab1 
               Height          =   2745
               Left            =   240
               TabIndex        =   205
               Top             =   120
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   4842
               _Version        =   393216
               Tabs            =   2
               Tab             =   1
               TabsPerRow      =   2
               TabHeight       =   420
               TabCaption(0)   =   "«·›—Ê⁄"
               TabPicture(0)   =   "FrmOvers.frx":72F4
               Tab(0).ControlEnabled=   0   'False
               Tab(0).Control(0)=   "Fra(1)"
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "‰ﬁ«ÿ «·»Ì⁄"
               TabPicture(1)   =   "FrmOvers.frx":7310
               Tab(1).ControlEnabled=   -1  'True
               Tab(1).Control(0)=   "Fra(9)"
               Tab(1).Control(0).Enabled=   0   'False
               Tab(1).ControlCount=   1
               Begin VB.Frame Fra 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  ‰ﬁ«ÿ «·»Ì⁄ «· Ì Ìÿ»ﬁ ›ÌÂ« «·⁄—÷"
                  Height          =   2316
                  Index           =   9
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   215
                  Top             =   360
                  Width           =   6492
                  Begin VB.TextBox Text5 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   -1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   216
                     Top             =   600
                     Width           =   855
                  End
                  Begin XtremeSuiteControls.CheckBox chkAllPos 
                     Height          =   252
                     Left            =   4800
                     TabIndex        =   217
                     Top             =   240
                     Width           =   1452
                     _Version        =   786432
                     _ExtentX        =   2561
                     _ExtentY        =   444
                     _StockProps     =   79
                     Caption         =   "ﬂ· ‰ﬁ«ÿ «·»Ì⁄"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo dcPOS 
                     Bindings        =   "FrmOvers.frx":732C
                     Height          =   288
                     Left            =   1680
                     TabIndex        =   218
                     Top             =   240
                     Width           =   2172
                     _ExtentX        =   3836
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
                  Begin VSFlex8Ctl.VSFlexGrid grdPos 
                     Height          =   1152
                     Left            =   120
                     TabIndex        =   219
                     Top             =   720
                     Width           =   6228
                     _cx             =   10985
                     _cy             =   2032
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
                     Rows            =   1
                     Cols            =   6
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmOvers.frx":7341
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
                  Begin ImpulseButton.ISButton CmddelPos 
                     Height          =   276
                     Index           =   0
                     Left            =   5640
                     TabIndex        =   220
                     Top             =   1920
                     Width           =   696
                     _ExtentX        =   1217
                     _ExtentY        =   476
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
                     ButtonImage     =   "FrmOvers.frx":741F
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton CmddelPos 
                     Height          =   396
                     Index           =   1
                     Left            =   120
                     TabIndex        =   221
                     Top             =   1800
                     Width           =   1176
                     _ExtentX        =   2064
                     _ExtentY        =   688
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "Õ–› «·ﬂ·"
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
                     ButtonImage     =   "FrmOvers.frx":79B9
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton cmdInsertPos 
                     Height          =   390
                     Left            =   390
                     TabIndex        =   222
                     Top             =   180
                     Width           =   720
                     _ExtentX        =   1270
                     _ExtentY        =   688
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "≈œ—«Ã"
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
                     ButtonImage     =   "FrmOvers.frx":7F53
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "‰ﬁÿÂ »Ì⁄ „ÕœœÂ"
                     Height          =   192
                     Index           =   59
                     Left            =   3852
                     TabIndex        =   223
                     Top             =   240
                     Width           =   876
                  End
               End
               Begin VB.Frame Fra 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  «·›—Ê⁄ «· Ì Ìÿ»ﬁ ›ÌÂ« «·⁄—÷"
                  Height          =   2316
                  Index           =   1
                  Left            =   -74760
                  RightToLeft     =   -1  'True
                  TabIndex        =   206
                  Top             =   360
                  Width           =   6492
                  Begin VB.TextBox Text2 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   -1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   207
                     Top             =   600
                     Width           =   855
                  End
                  Begin XtremeSuiteControls.CheckBox ChAllBranch 
                     Height          =   252
                     Left            =   4920
                     TabIndex        =   208
                     Top             =   240
                     Width           =   1332
                     _Version        =   786432
                     _ExtentX        =   2350
                     _ExtentY        =   444
                     _StockProps     =   79
                     Caption         =   "ﬂ· «·›—Ê⁄"
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcbBranch1 
                     Bindings        =   "FrmOvers.frx":82ED
                     Height          =   288
                     Left            =   1680
                     TabIndex        =   209
                     Top             =   240
                     Width           =   2172
                     _ExtentX        =   3836
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
                  Begin VSFlex8Ctl.VSFlexGrid FgBranch 
                     Height          =   1152
                     Left            =   120
                     TabIndex        =   210
                     Top             =   720
                     Width           =   6228
                     _cx             =   10985
                     _cy             =   2032
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
                     Rows            =   1
                     Cols            =   6
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmOvers.frx":8302
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
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   276
                     Index           =   13
                     Left            =   5640
                     TabIndex        =   211
                     Top             =   1920
                     Width           =   696
                     _ExtentX        =   1217
                     _ExtentY        =   476
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
                     ButtonImage     =   "FrmOvers.frx":83DC
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton Cmd 
                     Height          =   396
                     Index           =   15
                     Left            =   120
                     TabIndex        =   212
                     Top             =   1800
                     Width           =   1176
                     _ExtentX        =   2064
                     _ExtentY        =   688
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "Õ–› «·ﬂ·"
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
                     ButtonImage     =   "FrmOvers.frx":8976
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton BtonAdd 
                     Height          =   390
                     Left            =   390
                     TabIndex        =   213
                     Top             =   180
                     Width           =   720
                     _ExtentX        =   1270
                     _ExtentY        =   688
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "≈œ—«Ã"
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
                     ButtonImage     =   "FrmOvers.frx":8F10
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "›—⁄ „Õœœ"
                     Height          =   288
                     Index           =   17
                     Left            =   3360
                     TabIndex        =   214
                     Top             =   240
                     Width           =   1368
                  End
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Height          =   600
               Index           =   7
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   4530
               Width           =   17955
               Begin VB.TextBox TxtCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   14640
                  TabIndex        =   153
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   1575
               End
               Begin VB.TextBox TxtPriceDit 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   5400
                  TabIndex        =   143
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   1095
               End
               Begin VB.ComboBox DcbTypePoliceyDit 
                  Height          =   315
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   142
                  Top             =   180
                  Width           =   1455
               End
               Begin VB.TextBox TxtAmountDit 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   7200
                  TabIndex        =   141
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   1095
               End
               Begin VB.TextBox Text24 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   -1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   140
                  Top             =   600
                  Width           =   855
               End
               Begin VB.TextBox TxtRateD 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1080
                  TabIndex        =   139
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   975
               End
               Begin MSDataListLib.DataCombo DcbItemDit 
                  Bindings        =   "FrmOvers.frx":92AA
                  Height          =   315
                  Left            =   10680
                  TabIndex        =   144
                  Top             =   180
                  Width           =   3855
                  _ExtentX        =   6800
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
               Begin MSDataListLib.DataCombo DcbUnitDit 
                  Bindings        =   "FrmOvers.frx":92BF
                  Height          =   315
                  Left            =   9120
                  TabIndex        =   145
                  Top             =   180
                  Width           =   1335
                  _ExtentX        =   2355
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
               Begin ImpulseButton.ISButton BtonAdd3 
                  Height          =   390
                  Left            =   90
                  TabIndex        =   146
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈œ«—Ã"
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
                  ButtonImage     =   "FrmOvers.frx":92D4
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·‰”»…"
                  Height          =   285
                  Index           =   46
                  Left            =   2040
                  TabIndex        =   152
                  Top             =   180
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·Œ’„"
                  Height          =   285
                  Index           =   47
                  Left            =   3840
                  TabIndex        =   151
                  Top             =   180
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄—"
                  Height          =   285
                  Index           =   48
                  Left            =   6360
                  TabIndex        =   150
                  Top             =   180
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   285
                  Index           =   49
                  Left            =   7440
                  TabIndex        =   149
                  Top             =   180
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊÕœ…"
                  Height          =   285
                  Index           =   50
                  Left            =   10800
                  TabIndex        =   148
                  Top             =   300
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·’‰›"
                  Height          =   285
                  Index           =   51
                  Left            =   15600
                  TabIndex        =   147
                  Top             =   180
                  Width           =   1365
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   "”Ì«”… «·Œ’„"
               Height          =   1185
               Index           =   3
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   3345
               Width           =   7248
               Begin VB.Frame Frame1 
                  Caption         =   "⁄—÷ „Œ’’"
                  Height          =   975
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   3975
                  Begin VB.ComboBox CboFromPrice 
                     Height          =   315
                     ItemData        =   "FrmOvers.frx":966E
                     Left            =   0
                     List            =   "FrmOvers.frx":9670
                     RightToLeft     =   -1  'True
                     TabIndex        =   162
                     Top             =   600
                     Width           =   1695
                  End
                  Begin VB.TextBox TxtDiscount 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   2520
                     RightToLeft     =   -1  'True
                     TabIndex        =   160
                     Top             =   600
                     Width           =   495
                  End
                  Begin VB.TextBox TxtGetFree 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   1200
                     RightToLeft     =   -1  'True
                     TabIndex        =   157
                     Top             =   240
                     Width           =   495
                  End
                  Begin VB.TextBox TxtSales 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   2520
                     RightToLeft     =   -1  'True
                     TabIndex        =   156
                     Top             =   240
                     Width           =   495
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "„‰ "
                     Height          =   285
                     Index           =   34
                     Left            =   1800
                     TabIndex        =   161
                     Top             =   600
                     Width           =   645
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "»‰”»… Œ’„"
                     Height          =   285
                     Index           =   33
                     Left            =   3000
                     TabIndex        =   159
                     Top             =   600
                     Width           =   765
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«Õ’· ⁄·Ì"
                     Height          =   285
                     Index           =   32
                     Left            =   1680
                     TabIndex        =   158
                     Top             =   240
                     Width           =   765
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«‘ —Ì"
                     Height          =   285
                     Index           =   31
                     Left            =   3120
                     TabIndex        =   155
                     Top             =   240
                     Width           =   645
                  End
               End
               Begin VB.ComboBox DcbtypPolicep 
                  Height          =   315
                  ItemData        =   "FrmOvers.frx":9672
                  Left            =   3960
                  List            =   "FrmOvers.frx":9674
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
                  Top             =   240
                  Width           =   1935
               End
               Begin VB.Frame Fra 
                  BackColor       =   &H00E2E9E9&
                  Height          =   1065
                  Index           =   4
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   2280
                  Width           =   6525
                  Begin VB.TextBox Text9 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   -1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   122
                     Top             =   600
                     Width           =   855
                  End
                  Begin VB.TextBox TxtPriceBisc1 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   120
                     TabIndex        =   121
                     TabStop         =   0   'False
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   2055
                  End
                  Begin VB.TextBox TxtAmountBisc1 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   120
                     TabIndex        =   120
                     TabStop         =   0   'False
                     Top             =   240
                     Width           =   5055
                  End
                  Begin VB.TextBox TxtPriceDis 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   120
                     TabIndex        =   119
                     TabStop         =   0   'False
                     Top             =   720
                     Visible         =   0   'False
                     Width           =   855
                  End
                  Begin VB.TextBox TxtAmountDis 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   120
                     TabIndex        =   118
                     TabStop         =   0   'False
                     Top             =   720
                     Width           =   2055
                  End
                  Begin MSDataListLib.DataCombo dcbUnitBisc1 
                     Bindings        =   "FrmOvers.frx":9676
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   123
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   975
                     _ExtentX        =   1720
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
                  Begin MSDataListLib.DataCombo DcbItemBisc1 
                     Bindings        =   "FrmOvers.frx":968B
                     Height          =   315
                     Left            =   4200
                     TabIndex        =   124
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1215
                     _ExtentX        =   2143
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
                  Begin MSDataListLib.DataCombo dcbUnitDis 
                     Bindings        =   "FrmOvers.frx":96A0
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   125
                     Top             =   720
                     Width           =   975
                     _ExtentX        =   1720
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
                  Begin MSDataListLib.DataCombo DcbItemDis 
                     Bindings        =   "FrmOvers.frx":96B5
                     Height          =   315
                     Left            =   4200
                     TabIndex        =   126
                     Top             =   720
                     Width           =   1215
                     _ExtentX        =   2143
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "’‰› «”«”Ì"
                     Height          =   285
                     Index           =   12
                     Left            =   4920
                     TabIndex        =   134
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1365
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÊÕœ…"
                     Height          =   285
                     Index           =   19
                     Left            =   3600
                     TabIndex        =   133
                     Top             =   720
                     Width           =   525
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ„Ì…"
                     Height          =   285
                     Index           =   20
                     Left            =   2040
                     TabIndex        =   132
                     Top             =   720
                     Width           =   525
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄—"
                     Height          =   285
                     Index           =   21
                     Left            =   2040
                     TabIndex        =   131
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   525
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "’‰› «·Œ’„"
                     Height          =   285
                     Index           =   22
                     Left            =   5520
                     TabIndex        =   130
                     Top             =   720
                     Width           =   885
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÊÕœ…"
                     Height          =   285
                     Index           =   23
                     Left            =   3600
                     TabIndex        =   129
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   525
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ„Ì…"
                     Height          =   285
                     Index           =   24
                     Left            =   5880
                     TabIndex        =   128
                     Top             =   240
                     Width           =   525
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄—"
                     Height          =   285
                     Index           =   26
                     Left            =   840
                     TabIndex        =   127
                     Top             =   720
                     Visible         =   0   'False
                     Width           =   525
                  End
               End
               Begin VB.Frame Fra 
                  BackColor       =   &H00E2E9E9&
                  Height          =   660
                  Index           =   5
                  Left            =   840
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   480
                  Width           =   5925
                  Begin VB.TextBox Text14 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   -1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   113
                     Top             =   600
                     Width           =   855
                  End
                  Begin VB.TextBox TxtRate 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   4320
                     TabIndex        =   112
                     TabStop         =   0   'False
                     Top             =   240
                     Width           =   1095
                  End
                  Begin VB.TextBox TxtAmountBisc2 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   4320
                     TabIndex        =   111
                     TabStop         =   0   'False
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1095
                  End
                  Begin VB.TextBox TxtPriceBisc2 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   2160
                     TabIndex        =   110
                     TabStop         =   0   'False
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”»…"
                     Height          =   285
                     Index           =   29
                     Left            =   5280
                     TabIndex        =   116
                     Top             =   240
                     Width           =   525
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ„Ì…"
                     Height          =   285
                     Index           =   11
                     Left            =   5280
                     TabIndex        =   115
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   525
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄—"
                     Height          =   285
                     Index           =   16
                     Left            =   3480
                     TabIndex        =   114
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   525
                  End
               End
               Begin VB.Frame Fra 
                  BackColor       =   &H00E2E9E9&
                  Height          =   675
                  Index           =   8
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   3120
                  Width           =   6525
                  Begin VB.TextBox TxtAmountDDis 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1440
                     TabIndex        =   102
                     TabStop         =   0   'False
                     Top             =   240
                     Width           =   735
                  End
                  Begin VB.TextBox TxtPriceDDis 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   120
                     TabIndex        =   101
                     TabStop         =   0   'False
                     Top             =   240
                     Width           =   855
                  End
                  Begin VB.TextBox Text17 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   -1440
                     RightToLeft     =   -1  'True
                     TabIndex        =   100
                     Top             =   600
                     Width           =   855
                  End
                  Begin MSDataListLib.DataCombo DcbUnitDDis 
                     Bindings        =   "FrmOvers.frx":96CA
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   103
                     Top             =   240
                     Width           =   975
                     _ExtentX        =   1720
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
                  Begin MSDataListLib.DataCombo DcbItemDDis 
                     Bindings        =   "FrmOvers.frx":96DF
                     Height          =   315
                     Left            =   4200
                     TabIndex        =   104
                     Top             =   240
                     Width           =   1215
                     _ExtentX        =   2143
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”⁄—"
                     Height          =   285
                     Index           =   45
                     Left            =   840
                     TabIndex        =   108
                     Top             =   240
                     Width           =   525
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "’‰› «·Œ’„"
                     Height          =   285
                     Index           =   42
                     Left            =   5520
                     TabIndex        =   107
                     Top             =   240
                     Width           =   885
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ„Ì…"
                     Height          =   285
                     Index           =   40
                     Left            =   2040
                     TabIndex        =   106
                     Top             =   240
                     Width           =   525
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÊÕœ…"
                     Height          =   285
                     Index           =   39
                     Left            =   3600
                     TabIndex        =   105
                     Top             =   240
                     Width           =   525
                  End
               End
               Begin ImpulseButton.ISButton BtonAdd2 
                  Height          =   420
                  Left            =   240
                  TabIndex        =   136
                  Top             =   -600
                  Visible         =   0   'False
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   741
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈œ—«Ã"
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
                  ButtonImage     =   "FrmOvers.frx":96F4
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰Ê⁄ «·Œ’„"
                  Height          =   285
                  Index           =   28
                  Left            =   5400
                  TabIndex        =   137
                  Top             =   240
                  Width           =   1365
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Height          =   4545
               Index           =   11
               Left            =   7464
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   60
               Width           =   10590
               Begin VB.TextBox txtPeriod2 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   870
                  TabIndex        =   197
                  TabStop         =   0   'False
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.Frame Fra 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ÿ»Ìﬁ «·⁄—÷ €·Ì"
                  Height          =   3600
                  Index           =   6
                  Left            =   -120
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   840
                  Width           =   10425
                  Begin VB.Frame Frame3 
                     Caption         =   "„ Ê”ÿ „»Ì⁄«  «·ÌÊ„"
                     Height          =   3135
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   173
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   4875
                     Begin VB.TextBox txtTotalPer 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   2265
                        TabIndex        =   200
                        TabStop         =   0   'False
                        Text            =   "9"
                        Top             =   930
                        Width           =   1095
                     End
                     Begin VB.TextBox txtResultValue 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   2730
                        TabIndex        =   194
                        TabStop         =   0   'False
                        Top             =   2640
                        Width           =   1095
                     End
                     Begin VB.TextBox txtTotalQtyP 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   2265
                        TabIndex        =   193
                        TabStop         =   0   'False
                        Text            =   "600"
                        Top             =   1290
                        Width           =   1095
                     End
                     Begin VB.TextBox txtAvgQtyD 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   2250
                        TabIndex        =   192
                        TabStop         =   0   'False
                        Text            =   "66.66"
                        Top             =   1665
                        Width           =   1110
                     End
                     Begin VB.TextBox txtProdArrive 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   120
                        TabIndex        =   184
                        TabStop         =   0   'False
                        Text            =   "4"
                        Top             =   555
                        Visible         =   0   'False
                        Width           =   795
                     End
                     Begin VB.TextBox txtMin 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   540
                        TabIndex        =   182
                        TabStop         =   0   'False
                        Top             =   2115
                        Width           =   1095
                     End
                     Begin VB.TextBox txtPlus 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   3180
                        TabIndex        =   180
                        TabStop         =   0   'False
                        Top             =   2085
                        Width           =   1095
                     End
                     Begin VB.TextBox txtPeriod 
                        Alignment       =   2  'Center
                        Height          =   315
                        Left            =   2745
                        TabIndex        =   177
                        TabStop         =   0   'False
                        Text            =   "5"
                        Top             =   555
                        Width           =   615
                     End
                     Begin MSComCtl2.DTPicker Startdate2 
                        Height          =   315
                        Left            =   2310
                        TabIndex        =   174
                        Top             =   225
                        Width           =   1455
                        _ExtentX        =   2566
                        _ExtentY        =   556
                        _Version        =   393216
                        Format          =   245366785
                        CurrentDate     =   45292
                     End
                     Begin MSComCtl2.DTPicker enddate2 
                        Height          =   315
                        Left            =   840
                        TabIndex        =   176
                        Top             =   240
                        Width           =   1335
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _Version        =   393216
                        Format          =   245366785
                        CurrentDate     =   41640
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "„ Ê”ÿ „»Ì⁄«  «·ÌÊ„+-«·“Ì«œ… Ê«·‰ﬁ’«‰*„œ… «·‘—«¡ «·„” Âœ›…"
                        ForeColor       =   &H000000FF&
                        Height          =   405
                        Index           =   58
                        Left            =   -180
                        TabIndex        =   204
                        Top             =   2640
                        Width           =   2505
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«Ã„«·Ì «·ﬂ„Ì… / «Ã„«·Ì «·„œ…"
                        ForeColor       =   &H000000FF&
                        Height          =   405
                        Index           =   57
                        Left            =   -360
                        TabIndex        =   203
                        Top             =   1740
                        Width           =   2505
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "’«›Ï «·„»Ì⁄«  -«·—’Ìœ"
                        ForeColor       =   &H000000FF&
                        Height          =   405
                        Index           =   56
                        Left            =   -360
                        TabIndex        =   202
                        Top             =   1350
                        Width           =   2505
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   " „œ… ﬁÌ«” «·„»Ì⁄«  + „œ… Ê’Ê· «·»÷«⁄…"
                        ForeColor       =   &H000000FF&
                        Height          =   405
                        Index           =   55
                        Left            =   -300
                        TabIndex        =   201
                        Top             =   930
                        Width           =   2505
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«Ã„«·Ì «·„œ…"
                        Height          =   285
                        Index           =   54
                        Left            =   3300
                        TabIndex        =   199
                        Top             =   960
                        Width           =   1515
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·ﬂ„ÌÂ «·‰Â«∆ÌÂ"
                        Height          =   225
                        Index           =   53
                        Left            =   3330
                        TabIndex        =   191
                        Top             =   2715
                        Width           =   1515
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«Ã„«·Ï ﬂ„ÌÂ «·„»Ì⁄« "
                        Height          =   285
                        Index           =   52
                        Left            =   3330
                        TabIndex        =   190
                        Top             =   1320
                        Width           =   1515
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "„ Ê”ÿ „»Ì⁄«  «·ÌÊ„"
                        Height          =   285
                        Index           =   44
                        Left            =   3330
                        TabIndex        =   189
                        Top             =   1665
                        Width           =   1515
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "„œ… Ê’Ê· «·»÷«⁄…"
                        Height          =   285
                        Index           =   43
                        Left            =   960
                        TabIndex        =   185
                        Top             =   615
                        Visible         =   0   'False
                        Width           =   1275
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·‰ﬁ’«‰"
                        Height          =   285
                        Index           =   41
                        Left            =   1380
                        TabIndex        =   183
                        Top             =   2115
                        Width           =   1515
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·“Ì«œ…"
                        Height          =   285
                        Index           =   38
                        Left            =   3330
                        TabIndex        =   181
                        Top             =   2100
                        Width           =   1515
                     End
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        Caption         =   "ÌÊ„"
                        Height          =   345
                        Left            =   2040
                        RightToLeft     =   -1  'True
                        TabIndex        =   179
                        Top             =   600
                        Width           =   645
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "„œ… ﬁÌ«” «·„»Ì⁄« "
                        Height          =   285
                        Index           =   37
                        Left            =   3330
                        TabIndex        =   178
                        Top             =   570
                        Width           =   1515
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«· «—ÌŒ"
                        Height          =   285
                        Index           =   36
                        Left            =   3480
                        TabIndex        =   175
                        Top             =   255
                        Width           =   1365
                     End
                  End
                  Begin VB.Frame Frame2 
                     Height          =   585
                     Left            =   5130
                     RightToLeft     =   -1  'True
                     TabIndex        =   170
                     Top             =   2970
                     Width           =   5175
                     Begin MSDataListLib.DataCombo DBCboClientName 
                        Height          =   315
                        Left            =   1320
                        TabIndex        =   171
                        Top             =   150
                        Width           =   2955
                        _ExtentX        =   5212
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   " «·„Ê—œ"
                        Height          =   195
                        Index           =   35
                        Left            =   3960
                        RightToLeft     =   -1  'True
                        TabIndex        =   172
                        Top             =   120
                        Width           =   1095
                     End
                  End
                  Begin VB.Frame Frame11 
                     Height          =   1635
                     Left            =   5100
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   720
                     Width           =   5175
                     Begin VB.ListBox ListGroupSelected 
                        Height          =   840
                        ItemData        =   "FrmOvers.frx":9A8E
                        Left            =   120
                        List            =   "FrmOvers.frx":9A95
                        RightToLeft     =   -1  'True
                        TabIndex        =   95
                        Top             =   240
                        Width           =   2205
                     End
                     Begin VB.ListBox ListGroupAll 
                        Height          =   840
                        ItemData        =   "FrmOvers.frx":9AAC
                        Left            =   2760
                        List            =   "FrmOvers.frx":9AB3
                        RightToLeft     =   -1  'True
                        TabIndex        =   94
                        Top             =   240
                        Width           =   2325
                     End
                     Begin VB.Label Label8 
                        Alignment       =   2  'Center
                        Caption         =   ">"
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   13.5
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   2310
                        RightToLeft     =   -1  'True
                        TabIndex        =   89
                        Top             =   240
                        Width           =   495
                     End
                     Begin VB.Label Label7 
                        Alignment       =   2  'Center
                        Caption         =   ">>"
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   13.5
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00FF0000&
                        Height          =   255
                        Left            =   2310
                        RightToLeft     =   -1  'True
                        TabIndex        =   88
                        Top             =   480
                        Width           =   495
                     End
                     Begin VB.Label Label6 
                        Alignment       =   2  'Center
                        Caption         =   "<<"
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
                        Height          =   255
                        Left            =   2310
                        RightToLeft     =   -1  'True
                        TabIndex        =   87
                        Top             =   720
                        Width           =   495
                     End
                     Begin VB.Label Label5 
                        Alignment       =   2  'Center
                        Caption         =   "<"
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
                        Left            =   2310
                        RightToLeft     =   -1  'True
                        TabIndex        =   86
                        Top             =   960
                        Width           =   495
                     End
                  End
                  Begin VB.OptionButton XPOptShowType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂ· «·«’‰«›"
                     ForeColor       =   &H00FF0000&
                     Height          =   210
                     Index           =   0
                     Left            =   8640
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   240
                     Value           =   -1  'True
                     Width           =   1185
                  End
                  Begin VB.OptionButton XPOptShowType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "’‰› „Õœœ  ≈Œ «— «·’‰›"
                     ForeColor       =   &H000000FF&
                     Height          =   210
                     Index           =   2
                     Left            =   5250
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   540
                     Width           =   2265
                  End
                  Begin VB.OptionButton XPOptShowType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Ã„Ê⁄«  „Õœœ…"
                     ForeColor       =   &H00FF0000&
                     Height          =   210
                     Index           =   1
                     Left            =   5760
                     RightToLeft     =   -1  'True
                     TabIndex        =   81
                     Top             =   480
                     Width           =   4065
                  End
                  Begin VB.TextBox Text1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H000000FF&
                     Height          =   315
                     Left            =   6120
                     RightToLeft     =   -1  'True
                     TabIndex        =   80
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   1455
                  End
                  Begin MSDataListLib.DataCombo DcItem1 
                     Height          =   315
                     Left            =   2640
                     TabIndex        =   84
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   3495
                     _ExtentX        =   6165
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   255
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcbUnitGroup 
                     Bindings        =   "FrmOvers.frx":9AC5
                     Height          =   315
                     Left            =   -690
                     TabIndex        =   90
                     Top             =   -720
                     Visible         =   0   'False
                     Width           =   1815
                     _ExtentX        =   3201
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   255
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
                  Begin MSDataListLib.DataCombo DcbUnitG 
                     Bindings        =   "FrmOvers.frx":9ADA
                     Height          =   315
                     Left            =   -690
                     TabIndex        =   92
                     Top             =   -120
                     Visible         =   0   'False
                     Width           =   1815
                     _ExtentX        =   3201
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   255
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
                  Begin MSDataListLib.DataCombo DcbUnit 
                     Bindings        =   "FrmOvers.frx":9AEF
                     Height          =   315
                     Left            =   0
                     TabIndex        =   96
                     Top             =   -60
                     Visible         =   0   'False
                     Width           =   1815
                     _ExtentX        =   3201
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   255
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
                  Begin ImpulseButton.ISButton ISButton1 
                     Height          =   270
                     Left            =   5040
                     TabIndex        =   195
                     Top             =   2610
                     Width           =   1380
                     _ExtentX        =   2434
                     _ExtentY        =   476
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "≈œ«—Ã"
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
                     ButtonImage     =   "FrmOvers.frx":9B04
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÊÕœ…"
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Index           =   13
                     Left            =   1260
                     TabIndex        =   97
                     Top             =   -30
                     Visible         =   0   'False
                     Width           =   1365
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÊÕœ…"
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Index           =   18
                     Left            =   390
                     TabIndex        =   93
                     Top             =   -120
                     Visible         =   0   'False
                     Width           =   1365
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÊÕœ…"
                     ForeColor       =   &H000000FF&
                     Height          =   285
                     Index           =   14
                     Left            =   270
                     TabIndex        =   91
                     Top             =   -720
                     Visible         =   0   'False
                     Width           =   1365
                  End
               End
               Begin VB.TextBox TxtNameShow 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   150
                  Width           =   8895
               End
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   -1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   600
                  Width           =   855
               End
               Begin MSComCtl2.DTPicker enddate 
                  Height          =   315
                  Left            =   2070
                  TabIndex        =   55
                  Top             =   495
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   245366785
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker Startdate 
                  Height          =   315
                  Left            =   4230
                  TabIndex        =   56
                  Top             =   525
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   245366785
                  CurrentDate     =   45292
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  Caption         =   "ÌÊ„"
                  Height          =   345
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   198
                  Top             =   525
                  Width           =   795
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„œ… «·‘—«¡ «·„” Âœ›…"
                  Height          =   255
                  Left            =   7080
                  RightToLeft     =   -1  'True
                  TabIndex        =   196
                  Top             =   600
                  Width           =   1695
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì»œ√›Ì "
                  Height          =   285
                  Index           =   2
                  Left            =   4920
                  TabIndex        =   60
                  Top             =   555
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ "
                  Height          =   285
                  Index           =   9
                  Left            =   9000
                  TabIndex        =   59
                  Top             =   150
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì‰ ÂÌ ›Ì"
                  Height          =   285
                  Index           =   5
                  Left            =   3120
                  TabIndex        =   57
                  Top             =   525
                  Width           =   1005
               End
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   780
               Left            =   0
               TabIndex        =   47
               Top             =   7680
               Width           =   2550
               _ExtentX        =   4498
               _ExtentY        =   1376
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "«—”«· ··«⁄ „«œ"
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
               Height          =   615
               Index           =   8
               Left            =   0
               TabIndex        =   51
               Top             =   23070
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   1085
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–›"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmOvers.frx":9E9E
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   630
               Index           =   10
               Left            =   0
               TabIndex        =   52
               Top             =   -5700
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   1111
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–›"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmOvers.frx":A438
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   600
               Index           =   11
               Left            =   -120
               TabIndex        =   53
               Top             =   51855
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   1058
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–›"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmOvers.frx":A9D2
               DrawFocusRectangle=   0   'False
            End
            Begin XtremeSuiteControls.RadioButton RdPrivatePolice 
               Height          =   405
               Left            =   13995
               TabIndex        =   72
               Top             =   4665
               Width           =   3930
               _Version        =   786432
               _ExtentX        =   6921
               _ExtentY        =   720
               _StockProps     =   79
               Caption         =   "·«’‰«› „Õœœ…"
               ForeColor       =   16711680
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdAllPolice 
               Height          =   405
               Left            =   2415
               TabIndex        =   73
               Top             =   2895
               Width           =   4575
               _Version        =   786432
               _ExtentX        =   8070
               _ExtentY        =   714
               _StockProps     =   79
               Caption         =   "”Ì«”… «Ã„«·Ì… ·ﬂ· «·«’‰«› «·„Õœœ…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7680
            Index           =   9
            Left            =   15
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   15
            Width           =   18000
            _cx             =   31750
            _cy             =   13547
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
            Appearance      =   5
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   0
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
            Begin VB.TextBox Text8 
               Alignment       =   1  'Right Justify
               Height          =   5760
               Left            =   4800
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   1560
               Width           =   945
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   4065
               Left            =   5988
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   2085
               Width           =   1536
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   4065
               Index           =   67
               Left            =   3420
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   2085
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   3840
               Index           =   68
               Left            =   5745
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   2520
               Width           =   30
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4605
               Index           =   69
               Left            =   4185
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   2085
               Width           =   630
            End
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   17
      Left            =   3480
      TabIndex        =   78
      Top             =   9480
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "‰”Œ… „„«À·…"
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
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—ﬁ„"
      Height          =   285
      Index           =   3
      Left            =   16200
      TabIndex        =   50
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Â–… «·‘«‘…  ﬁÊ„ » ”ÃÌ· ÿ·» ”›… ‰ﬁœÌ… ÊÌ „ «Õ ”«» ﬁÌ„… «·œ›⁄ «·Ì«"
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
      Height          =   660
      Index           =   25
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3450
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   120
      Top             =   3720
      Width           =   6015
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   18090
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   285
      Index           =   4
      Left            =   11280
      TabIndex        =   24
      Top             =   600
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   13710
      TabIndex        =   23
      Top             =   615
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   16125
      TabIndex        =   22
      Top             =   9675
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2430
      TabIndex        =   21
      Top             =   9630
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   690
      TabIndex        =   20
      Top             =   9630
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   19
      Top             =   9540
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   18
      Top             =   9540
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   18870
      TabIndex        =   17
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "Frmovers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String
Public LngCol As Double
Public LngRow As Double
Public mIndex As Integer
Private Type ItemsOverData
        QtyDay As Double
        QtyPeriod As Double
        FinalQty As Double
End Type
 Public Ch As Boolean
 Dim Account_Code_dynamic As String
 Public Item As Integer

Private Function GetItemOverData(ItemID As Long) As ItemsOverData
    Dim ret As ItemsOverData
    Dim s   As String
    s = ""
    s = s & "SELECT Item_ID,  "
    s = s & "       SUM(   CASE "
    s = s & "                  WHEN Transaction_Type = 21 THEN "
    s = s & "                      Quantity "
    s = s & "                  WHEN Transaction_Type = 9 THEN "
    s = s & "                      -Quantity "
    s = s & "                  ELSE "
    s = s & "                      0 "
    s = s & "              END "
    s = s & "          ) Quantity  "
    '    s = s & "       ROUND(SUM(   CASE "
    '    s = s & "                        WHEN Transaction_Type = 21 THEN "
    '    s = s & "                            Quantity "
    '    s = s & "                        WHEN Transaction_Type = 9 THEN "
    '    s = s & "                            -Quantity "
    '    s = s & "                        ELSE "
    '    s = s & "                            0 "
    '    s = s & "                    END "
    '    s = s & "                ) / 30, "
    '    s = s & "             2 "
    '    s = s & "            ) Avrg "
    s = s & " FROM dbo.Transaction_Details "
    s = s & "    JOIN dbo.Transactions "
    s = s & "        ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID "
    s = s & "WHERE Transactions.Transaction_Date "
    s = s & "BETWEEN '" & Format(Startdate2, "yyyy-MM-dd") & "' AND '" & Format(enddate2, "yyyy-MM-dd") & "' "
    s = s & " And Item_ID = " & ItemID
    s = s & "GROUP BY Item_ID; "
    '/////////////////
    Dim ItemBalance As Double
    Dim RsBlncs     As ADODB.Recordset
    Set RsBlncs = GetItemQuantityStock(ItemID)
    If Not RsBlncs.EOF Then
        ItemBalance = val(RsBlncs!TotalQty & "")
    End If
    RsBlncs.Close
    '///////////////
    Dim rs As New ADODB.Recordset
    rs.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    Dim totalPer As Double
    
    If IsDate(Startdate2) And IsDate(enddate2) Then
        ' txtPeriod = DateDiff("d", Startdate2, enddate2)
        totalPer = DateDiff("d", Startdate2, enddate2) + val(txtProdArrive)
    Else
        totalPer = 1
    End If
    
    If rs.EOF Then
        ret.QtyDay = 0
        ret.QtyPeriod = 0
    Else
        ' ret.QtyDay = val(rs!Avrg & "")
        '’«›Ï «·„»Ì⁄«  -«·—’Ìœ
        ret.QtyPeriod = ItemBalance - val(rs!Quantity & "")
        '«Ã„«·Ì «·ﬂ„Ì… / «Ã„«·Ì «·„œ…
        ret.QtyDay = ret.QtyPeriod / IIf(totalPer <> 0, totalPer, 1)
    End If
    
    Dim Dayes As Integer, MainDayes As Integer
   
    Dayes = DateDiff("d", Startdate2, enddate2)  ' › —Â «·«„«‰
    MainDayes = DateDiff("d", StartDate, EndDate)  ' „œ… «·‘—«¡ «·„” Âœ›  ' › —Â «·„‘ —Ì« 
    
    '*******************************
    Dim PlusValue As Double
    Dim MinValue  As Double
    Dim purQty    As Double
    Dim FinalQty  As Double
    PlusValue = val(txtPlus)
    MinValue = val(txtMin)
    ' „ Ê”ÿ „»Ì⁄«  «·ÌÊ„+-«·“Ì«œ… Ê«·‰ﬁ’«‰*„œ… «·‘—«¡ «·„” Âœ›…
    'FinalQty = ItemBalance - ((ret.QtyDay * MainDayes) + (ret.QtyDay * Dayes)) + (PlusValue - MinValue)
    If PlusValue > 0 Then
        FinalQty = (ret.QtyDay + PlusValue) * MainDayes
    Else
        FinalQty = (ret.QtyDay - MinValue) * MainDayes
    End If
     
    ret.FinalQty = FinalQty
    '*******************************
    
    GetItemOverData = ret
End Function
'Private Sub Accredit_Click()
'    Dim BeginTrans As Boolean
'
'    Cn.BeginTrans
'    BeginTrans = True
'
'    If IsNull(rs("Posted")) Then
'        rs("Posted") = user_id
''        rs("PostedDate") = Time
'   Else
'       rs("Posted") = Null
'      rs("PostedDate") = Time
'   End If
'
'   rs.update
'If SystemOptions.UserInterface = ArabicInterface Then
'   Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
'Else
'Accredit.Caption = "Sent To approval "
'End If

  '  Cn.CommitTrans
 '   BeginTrans = False
'FillApprovedTable
'    Retrive (val(Me.XPTxtID.text))
'End Sub

Sub retrivePOS()
    Dim ID        As Integer
    Dim RsDetails As ADODB.Recordset
    Dim StrSQL    As String
    Dim i         As Integer
    If Me.chkAllPos.value = xtpChecked Then
        Set RsDetails = New ADODB.Recordset
        StrSQL = " SELECT BoxID, BoxName,BoxNamee FROM Tblposdata where 1=1"

        RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Else
        Set RsDetails = New ADODB.Recordset
        StrSQL = " SELECT BoxID, BoxName,BoxNamee FROM Tblposdata Where  BoxID = " & val(Me.dcPOS.BoundText) & ""

        RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    End If
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        With Me.grdPos
            ID = .rows
            .rows = .rows + RsDetails.RecordCount
            For i = ID To .rows - 1
     
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = RsDetails("BoxID").value
                .TextMatrix(i, .ColIndex("name")) = RsDetails("BoxName").value
                .TextMatrix(i, .ColIndex("POSId")) = RsDetails("BoxID").value
                RsDetails.MoveNext
            Next i
        End With
    End If
End Sub


Private Sub BtonAdd_Click()

retriveBranch
End Sub

Sub retriveBranch()
    Dim ID        As Integer
    Dim RsDetails As ADODB.Recordset
    Dim StrSQL    As String
    Dim i         As Integer
    Dim s As String
    If Me.ChAllBranch.value = xtpChecked Then
        Set RsDetails = New ADODB.Recordset
        StrSQL = " select * from TblBranchesData where 1=1"
        s = " Select Commonname,CSR,Privatekey,SerialNumber,SecretKey,PublickeycertPem,OrganizationName,Invoicetype,DefaultInvoicetype,"
        s = s & " Company_Comment,StreetName,AdditionalStreetName,BuildingNumber,PlotIdentification,CityName,PostalZone,branch_Code,branch_name,branch_id"
        s = s & " CountrySubentity,CitySubdivisionName,Company_Name_Eng,VATRegNo,Company_arabic_Name,industrey,SendingMode "
        s = s & " from TblBranchesData where 1 = 1"

StrSQL = s
        RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Else
        Set RsDetails = New ADODB.Recordset
           s = " Select Commonname,CSR,Privatekey,SerialNumber,SecretKey,PublickeycertPem,OrganizationName,Invoicetype,DefaultInvoicetype,"
        s = s & " Company_Comment,StreetName,AdditionalStreetName,BuildingNumber,PlotIdentification,CityName,PostalZone,branch_Code,branch_name,branch_id"
        s = s & " CountrySubentity,CitySubdivisionName,Company_Name_Eng,VATRegNo,Company_arabic_Name,industrey,SendingMode "
        s = s & " from TblBranchesData where branch_id = " & val(Me.DcbBranch1.BoundText) & ""


        
        StrSQL = s

        RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    End If
    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        With Me.FgBranch
            ID = .rows
            .rows = .rows + RsDetails.RecordCount

            For i = ID To .rows - 1
     
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = RsDetails("branch_Code").value
                .TextMatrix(i, .ColIndex("name")) = RsDetails("branch_name").value
                .TextMatrix(i, .ColIndex("branchid")) = RsDetails("branch_id").value
                ' .TextMatrix(i, .ColIndex("group")) = RsDetails("GroupName").value
                RsDetails.MoveNext
            Next i
        End With
    End If
End Sub
'Sub retriveItemDitailsGroup()
'Dim id As Integer
'
'Dim i As Integer
'
'       Dim k As Integer
'        With Me.FgItemPloice
'       ' id = .Rows
'       '.Rows = .Rows + FgItems.Rows - 1
'

'        For i = 1 To .Rows - 1
'
    
'          .TextMatrix(i, .ColIndex("Ser")) = i
          ' .TextMatrix(i, .ColIndex("id")) = RsDetails("Fullcode").value
           '.TextMatrix(i, .ColIndex("name")) = FgItems.TextMatrix(k, FgItems.ColIndex("name"))
           ' .TextMatrix(i, .ColIndex("unite")) = FgItems.TextMatrix(k, FgItems.ColIndex("unite"))
          '.'TextMatrix(i, .ColIndex("group")) = FgItems.TextMatrix(k, FgItems.ColIndex("group"))
          '.TextMatrix(i, .ColIndex("untgroup")) = FgItems.TextMatrix(k, FgItems.ColIndex("unite1"))
           
          ' .TextMatrix(i, .ColIndex("typedisid")) = Me.DcbtypPolicep.ListIndex
'            .TextMatrix(i, .ColIndex("typedis")) = Me.DcbtypPolicep.ListIndex
'          .TextMatrix(i, .ColIndex("discount")) = Me.TxtRate.text
'          .TextMatrix(i, .ColIndex("amountdis")) = Me.TxtAmountDis.text
'          .TextMatrix(i, .ColIndex("pricedis")) = Me.TxtPriceDis.text
'              If Me.DcbtypPolicep.ListIndex < 2 Then
'              .TextMatrix(i, .ColIndex("itemdisid")) = ""
'               .TextMatrix(i, .ColIndex("unitdisid")) = ""
'          .TextMatrix(i, .ColIndex("unitdis")) = ""
'          .TextMatrix(i, .ColIndex("itemdis")) = ""
'          .TextMatrix(i, .ColIndex("unitdis")) = ""
'          Else
          ' .TextMatrix(i, .ColIndex("amountdis")) = ""
'          .TextMatrix(i, .ColIndex("discount")) = ""
'          End If
'           If Me.DcbtypPolicep.ListIndex = 2 Then
'              .TextMatrix(i, .ColIndex("itemdisid")) = FgItemPloice.TextMatrix(i, FgItemPloice.ColIndex("id"))
'               .TextMatrix(i, .ColIndex("unitdisid")) = FgItemPloice.TextMatrix(i, FgItemPloice.ColIndex("uniteId"))
''          .TextMatrix(i, .ColIndex("unitdis")) = FgItemPloice.TextMatrix(i, FgItemPloice.ColIndex("unite"))
 '         .TextMatrix(i, .ColIndex("itemdis")) = FgItemPloice.TextMatrix(i, FgItemPloice.ColIndex("name"))
 '         .TextMatrix(i, .ColIndex("unitdis")) = FgItemPloice.TextMatrix(i, FgItemPloice.ColIndex("unite"))
 '         Else
 '         .TextMatrix(i, .ColIndex("unitdisid")) = val(Me.dcbUnitDis.BoundText)
 '         .TextMatrix(i, .ColIndex("itemdisid")) = val(Me.DcbItemDis.BoundText)
 '         .TextMatrix(i, .ColIndex("itemdis")) = Me.DcbItemDis.text
 '         .TextMatrix(i, .ColIndex("unitdis")) = Me.dcbUnitDis.text
 '         End If
          
          
 '
'         If Me.DcbtypPolicep.ListIndex = 0 Or Me.DcbtypPolicep.ListIndex = 1 Then
'.TextMatrix(i, .ColIndex("amount")) = Me.TxtAmountBisc2.text
'           .TextMatrix(i, .ColIndex("price")) = Me.TxtPriceBisc2.text

'Else
'
'.TextMatrix(i, .ColIndex("amount")) = Me.TxtAmountBisc1.text
'           .TextMatrix(i, .ColIndex("price")) = Me.TxtPriceBisc1.text
'End If

'
'        Next i
'End With
    
'End Sub
'sa Sub retriveItemDitails()
'sa Dim id As Integer
'FgItemPloice.ColHidden(4) = True
'FgItemPloice.ColHidden(5) = True
'sa Dim i As Integer

       
  'sa      With Me.FgItemPloice
  'sa      id = .Rows
  'sa     .Rows = .Rows + 1


    'sa    For i = id To .Rows - 1
     
   'sa       .TextMatrix(i, .ColIndex("Ser")) = i
          
     'sa     .TextMatrix(i, .ColIndex("id")) = Me.DcbItemDit.BoundText
     'sa      .TextMatrix(i, .ColIndex("name")) = Me.DcbItemDit.text
     'sa       .TextMatrix(i, .ColIndex("unite")) = Me.DcbUnitDit.text
     'sa       .TextMatrix(i, .ColIndex("uniteid")) = Me.DcbUnitDit.BoundText
     'sa     .TextMatrix(i, .ColIndex("price")) = Me.TxtPriceDit.text
    'sa      .TextMatrix(i, .ColIndex("amount")) = Me.TxtAmountDit.text
    'sa       If Me.DcbTypePoliceyDit.ListIndex = 2 Then
     'sa      DcbItemDDis.BoundText = DcbItemDit.BoundText
     'sa      End If
           '.TextMatrix(i, .ColIndex("typedisid")) = Me.DcbTypePoliceyDit.ListIndex
        'sa   .TextMatrix(i, .ColIndex("typedis")) = Me.DcbTypePoliceyDit.ListIndex
       'sa    .TextMatrix(i, .ColIndex("discount")) = Me.TxtRateD.text
          ' .TextMatrix(i, .ColIndex("itemdisid")) = Me.DcbItemDDis.BoundText
          '  .TextMatrix(i, .ColIndex("itemdis")) = Me.DcbItemDDis.text
          '.TextMatrix(i, .ColIndex("unitdis")) = Me.DcbUnitDDis.text
          '.TextMatrix(i, .ColIndex("unitdisid")) = Me.DcbUnitDDis.BoundText
        'sa  .TextMatrix(i, .ColIndex("amountdis")) = Me.TxtAmountDDis.text
       'sa   .TextMatrix(i, .ColIndex("pricedis")) = Me.TxtPriceDDis.text
          '  If Me.DcbtypPolicep.ListIndex < 2 Then
          '    .TextMatrix(i, .ColIndex("itemdisid")) = ""
          '     .TextMatrix(i, .ColIndex("unitdisid")) = ""
          '.TextMatrix(i, .ColIndex("unitdis")) = ""
          '.TextMatrix(i, .ColIndex("itemdis")) = ""
          '.TextMatrix(i, .ColIndex("unitdis")) = ""
          'Else
           
          '.TextMatrix(i, .ColIndex("discount")) = ""
          'End If
           
 'sa       Next i
'saEnd With
    
'sa End Sub
'Sub retriveItemGroup()
'Dim id As Integer
'Dim RsDetails As ADODB.Recordset
'Dim StrSQL As String
'Dim i As Integer
' Set RsDetails = New ADODB.Recordset
'If Me.DcbGroup.BoundText <> "" Then
' If Me.DcbItem.BoundText <> "" Then
'
'StrSQL = " select * from TblItems where GroupID=" & Me.DcbGroup.BoundText & " and ItemID =" & Me.DcbItem.BoundText & ""
'Else
'StrSQL = " select * from TblItems where GroupID=" & Me.DcbGroup.BoundText & " "
'    End If
'Else
'If Me.DcbItem.BoundText <> "" Then
'StrSQL = " select * from TblItems where ItemID= " & Me.DcbItem.BoundText & ""
'End If
'End If
'RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'  If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        With Me.FgItems
'        id = .Rows
'       .Rows = .Rows + RsDetails.RecordCount
'
'
'        For i = id To .Rows - 1
'
'          .TextMatrix(i, .ColIndex("Ser")) = i
'           .TextMatrix(i, .ColIndex("id")) = RsDetails("Fullcode").value
'           .TextMatrix(i, .ColIndex("name")) = RsDetails("ItemName").value
'            .TextMatrix(i, .ColIndex("unite")) = Me.DcbUnit.text
'          .TextMatrix(i, .ColIndex("group")) = Me.DcbGroup.text
'          .TextMatrix(i, .ColIndex("unite1")) = Me.DcbUnitGroup.text
'            RsDetails.MoveNext
'        Next i
'End With
'    End If
'End Sub
'Sub retriveItem()
'Dim id As Integer
'Dim RsDetails As ADODB.Recordset
'Dim StrSQL As String
'Dim i As Integer
' Set RsDetails = New ADODB.Recordset
'If Me.DcbGroup.BoundText <> "" Then
' If Me.DcbItem.BoundText <> "" Then
'
'StrSQL = " select * from TblItems where GroupID=" & Me.DcbGroup.BoundText & " and ItemID =" & Me.DcbItem.BoundText & ""
'Else
'StrSQL = " select * from TblItems where GroupID=" & Me.DcbGroup.BoundText & " "
'    End If
'Else
'If Me.DcbItem.BoundText <> "" Then
'StrSQL = " select * from TblItems where ItemID= " & Me.DcbItem.BoundText & ""
'End If
'End If
'RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'  If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        With Me.FgItems
'        id = .Rows
'       .Rows = .Rows + RsDetails.RecordCount


'        For i = id To .Rows - 1
'
'          .TextMatrix(i, .ColIndex("Ser")) = i
'           .TextMatrix(i, .ColIndex("id")) = RsDetails("Fullcode").value
'           .TextMatrix(i, .ColIndex("name")) = RsDetails("ItemName").value
'            .TextMatrix(i, .ColIndex("unite1")) = Me.DcbUnit.text
'          .TextMatrix(i, .ColIndex("group")) = Me.DcbGroup.text
'          .TextMatrix(i, .ColIndex("unite")) = Me.DcbUnit1.text
'            RsDetails.MoveNext
'        Next i
'End With
'    End If
'End Sub
Private Sub BtonAdd1_Click()

End Sub

Private Sub BtonAdd2_Click()
' retriveItemDitailsGroup
End Sub

Private Sub BtonAdd3_Click()
 'retriveItemDitails
End Sub

Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

            'ShowGL_cc Me.TxtNoteSerial.text, , 200, val(Me.TxtNoteID.text)

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
         
            Fra(0).Enabled = True
            FgBranch.Clear flexClearScrollable, flexClearEverything
            FgBranch.rows = 1
            Me.FgBranch.Enabled = True
            grdPos.Clear flexClearScrollable, flexClearEverything
            grdPos.rows = 1
            Me.grdPos.Enabled = True
            
            ListGroupSelected.Clear
            Me.FgItemPloice.Clear flexClearScrollable, flexClearEverything
            FgItemPloice.rows = 1
            Me.FgItemPloice.Enabled = True
            RdAllPolice.value = False
        
            Me.FgItems.Clear flexClearScrollable, flexClearEverything
            FgItems.rows = 1
            XPOptShowType(1).value = False
            Me.FgItems.Enabled = True
            RdPrivatePolice.value = True
            'Frame10.Enabled = True
            'XPOptShowType(1).value = False
            'XPOptShowType(0).value = True
            TxtModFlg.text = "N"
            clear_all Me
            
            If mIndex = 0 Then
                Option1(0).value = True
            ElseIf mIndex = 1 Then
                Option1(1).value = True
            End If

            Option1_Click 1
            'Me.DcbOrderStatus.ListIndex = 1
            '  XPOptShowType(0).value = True
            '     GRID2.Clear flexClearScrollable, flexClearEverything
            'GRID2.Rows = 1
            Me.DCboUserName.BoundText = user_id
            '  TxtPaymentCounts.text = 1
            Me.DcbBranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            
            Accredit.Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
            Else
                Accredit.Caption = " send to Approval   "
            End If
                                               
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            'Frame10.Enabled = True
            'Frame11.Enabled = True
            Me.DCboUserName.BoundText = user_id

        Case 2
    
            Dim Msg As String

            If Trim(Me.DcbBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.DcbBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.DcbBranch.BoundText

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
            Ch = False
            Load FrmSearchItemShow
            FrmSearchItemShow.show

        Case 6
            Unload Me

        Case 7
            'ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 21

            '  RemoveGridRowGr
        Case 14
            Me.FgItems.Clear flexClearScrollable, flexClearEverything
            FgItems.rows = 1
 
        Case 13
            RemoveGridRowBr

        Case 15
            Me.FgBranch.Clear flexClearScrollable, flexClearEverything
            FgBranch.rows = 1
 
        Case 12
            RemoveGridRowPolice

        Case 16
            Me.FgItemPloice.Clear flexClearScrollable, flexClearEverything
            FgItemPloice.rows = 1
 
        Case 17
 
            TxtModFlg.text = "N"
            Me.XPTxtID.text = ""

            'sa frmCashCustomerSearch.RetrunType = 3
            'sa frmCashCustomerSearch.TxtCopun = TxtNameShow.text
            'sa  frmCashCustomerSearch.doit
            'sa frmCashCustomerSearch.show vbModal
 
        Case 9
            On Error Resume Next
            Dim StrFileName As String
            StrFileName = App.path & "overs.xls"

            If Dir(StrFileName) <> "" Then
                Kill StrFileName
            End If

            'Grid.RightToLeft = True
            Me.FgItemPloice.saveGrid StrFileName, flexFileExcel, True
            OpenFile StrFileName
        
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmddelPos_Click(Index As Integer)
    If Index = 1 Then
        Me.grdPos.Clear flexClearScrollable, flexClearEverything
        grdPos.rows = 1
    ElseIf Index = 0 Then
        RemoveGridRowBr2
    End If
End Sub

'Function print_report(Optional NoteSerial As String)
'
'
'   Dim MySQL As String
''    Dim RsData As New ADODB.Recordset
'   Dim xApp As New CRAXDRT.Application
'   Dim xReport As CRAXDRT.Report
'   Dim CViewer As ClsReportViewer
'   Dim StrReportTitle As String
'   Dim StrFileName As String
'   Dim Msg As String
'MySQL = "SELECT     dbo.TblCommisReceDetails.ID_Aut, dbo.TblCommisReceDetails.DateOp, dbo.TblCommisReceDetails.Total, dbo.TblCommisReceDetails.Fitter, "
' MySQL = MySQL & "                     dbo.TblCommisReceDetails.Operation, dbo.TblCommisReceDetails.PerceTage, dbo.TblCommisReceDetails.PerceTageValue, dbo.TblCommisReceDetails.id2,"
' MySQL = MySQL & "                     dbo.TblCommisRece.id, dbo.TblCommisRece.FitterID, dbo.TblCommisRece.DateFrom, dbo.TblCommisRece.DateTo, dbo.TblCommisRece.RecordDate,"
'MySQL = MySQL & "                      dbo.TblCommisRece.AllFit, dbo.TblCommisRece.LimitFit, dbo.TblCommisRece.UserID, dbo.TblCommisReceDetails.id AS idd, dbo.TblCommisReceDetails.PriceFitter,"
'MySQL = MySQL & "                      dbo.TblCommisReceDetails.Emp_id , dbo.TblCommisReceDetails.plateno, dbo.TblCommisReceDetails.Type, dbo.TblCommisReceDetails.Model"
'MySQL = MySQL & " FROM         dbo.TblCommisRece INNER JOIN"
'MySQL = MySQL & "                      dbo.TblCommisReceDetails ON dbo.TblCommisRece.id = dbo.TblCommisReceDetails.id2"
'MySQL = MySQL & " Where (dbo.TblCommisRece.id =" & val(XPTxtID.text) & ")"
'MySQL = MySQL & " Where (dbo.TblCommisRece.id =" & val(XPTxtID.text) & ")"

 
 'MySQL = MySQL & "   Where (dbo.TblTreatment.id =" & val(XPTxtID.text) & ")"

' If SystemOptions.UserInterface = ArabicInterface Then
'          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCommisRece.rpt"
'     Else
'        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCommisRece.rpt"
'       End If
''    If Dir(StrFileName) = "" Then
 '       'GetMsgs 139, vbExclamation
 '       Screen.MousePointer = vbDefault
 '       Exit Function
 '   End If
'
'    Set RsData = New ADODB.Recordset
'    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
''        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
 '       ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
 '       StrReportTitle = "" '& StrAccountName
 '       'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
 '       '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
 '       'End If
 '       'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
 '       '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
 '   Else
 
 '       xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
 '       xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
 '       StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
 '   End If

 '   xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
 '   xReport.reporttitle = StrReportTitle
 '   xReport.EnableParameterPrompting = False
 '   xReport.ApplicationName = App.Title
 '   xReport.ReportAuthor = App.Title
 '   Set CViewer = New ClsReportViewer
 '   CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
'
'    RsData.Close
'    Set RsData = Nothing
'    Screen.MousePointer = vbDefault


 
  
 
'End Function

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub





Private Sub cmdInsertPos_Click()
retrivePOS
End Sub

'Private Sub DcbGroup_Change()
'Dim Dcombos As ClsDataCombos
'    Set Dcombos = New ClsDataCombos
'    If Me.DcbGroup.BoundText <> "" Then
'Dcombos.GetItemsNamesupdate DcbItem, , , , , Me.DcbGroup.BoundText
'Else
'Dcombos.GetItemsNamesupdate DcbItem, , , , , 0
'
'End If
'End Sub
'
'Private Sub DcbGroup_Click(Area As Integer)
'Dim Dcombos As ClsDataCombos
'    Set Dcombos = New ClsDataCombos
'    If Me.DcbGroup.BoundText <> "" Then
'Dcombos.GetItemsNamesupdate DcbItem, , , , , Me.DcbGroup.BoundText
'Else
'Dcombos.GetItemsNamesupdate DcbItem, , , , , 0
'
'End If
'End Sub









'Private Sub DcbGroup_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF3 Then
'FrmGroupSearch.RetrunType = 1
'Load FrmGroupSearch
'           ' FrmSearchGroup.show
'           FrmGroupSearch.show
'End If
'End Sub

Private Sub DcbItem_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Item = 1

Load FrmItemSearch
          
            FrmItemSearch.show
End If
End Sub

Private Sub DcbItemDDis_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Item = 3
Load FrmItemSearch
          
            FrmItemSearch.show
            
End If
End Sub

Private Sub DcbItemDis_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Item = 4
Load FrmItemSearch
          
            FrmItemSearch.show
            
End If
End Sub

Private Sub DcbItemDit_Change()
Dim UnitID As Long
    Dim UnitName As String
  Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsUnits·byitemid Me.DcbUnitDit, val(Me.DcbItemDit.BoundText)
  
    GetDefaultItemUnit val(Me.DcbItemDit.BoundText), UnitID, UnitName
    DcbUnitDit.text = UnitName
    DcbUnitDit.BoundText = UnitID
        Me.TxtCode.text = GetItemCode(val(Me.DcbItemDit.BoundText))
        
End Sub

Private Sub DcbItemDit_Click(Area As Integer)
DcbItemDit_Change
End Sub

Private Sub DcbItemDit_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
 
Load FrmItemSearch
          FrmItemSearch.RetrunType = 2711
            FrmItemSearch.show
            
End If

End Sub

Private Sub DcbTypePoliceyDit_Click()
If Me.DcbTypePoliceyDit.ListIndex = 0 Then
TxtRateD.Visible = True
'Me.Fra(8).Visible = True
Me.Fra(8).Visible = False
 If SystemOptions.UserInterface = EnglishInterface Then
  Me.lbl(46).Caption = "Value"
 Else
 Me.lbl(46).Caption = "ﬁÌ„…"
 End If
Else
If Me.DcbTypePoliceyDit.ListIndex = 1 Then
TxtRateD.Visible = True
'Me.Fra(5).Visible = True
Me.Fra(8).Visible = False
 If SystemOptions.UserInterface = EnglishInterface Then
  Me.lbl(46).Caption = "Rate"
 Else
 Me.lbl(46).Caption = "‰”»…"
 End If
Else
If Me.DcbtypPolicep.ListIndex = 2 Then
'Me.Fra(5).Visible = False
DcbItemDDis.BoundText = DcbItemDit.BoundText
Me.Fra(8).Visible = True
TxtRateD.Visible = False
DcbItemDDis.Enabled = False
Me.DcbUnitDDis.Enabled = False


Else
DcbItemDDis.Enabled = True
TxtRateD.Visible = False
Me.DcbUnitDDis.Enabled = True
'Me.Fra(5).Visible = False
Me.Fra(8).Visible = True
End If

End If
End If
End Sub

Private Sub DcbtypPolicep_Click()

If Me.DcbtypPolicep.ListIndex = 0 Then
Me.Fra(5).Visible = True
Me.Fra(4).Visible = False
 If SystemOptions.UserInterface = EnglishInterface Then
  Me.lbl(29).Caption = "Value"
 Else
 Me.lbl(29).Caption = "ﬁÌ„…"
 End If
Else
If Me.DcbtypPolicep.ListIndex = 1 Then
Me.Fra(5).Visible = True
Me.Fra(4).Visible = False
 If SystemOptions.UserInterface = EnglishInterface Then
  Me.lbl(29).Caption = "Rate"
 Else
 Me.lbl(29).Caption = "‰”»…"
 End If

Else
If Me.DcbtypPolicep.ListIndex = 2 Then
Me.Fra(5).Visible = False
Me.Fra(4).Visible = True
DcbItemDis.Enabled = False
Me.dcbUnitDis.Enabled = False


Else
DcbItemDis.Enabled = True

Me.dcbUnitDis.Enabled = True
Me.Fra(5).Visible = False
Me.Fra(4).Visible = True
End If
End If
End If
   If Me.DcbtypPolicep.ListIndex = 4 Then
   Frame1.Visible = True
   Else
   Frame1.Visible = False
   End If
   
   
End Sub




Private Sub DcbUnit_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Item = 2
Load FrmSearchUnits
            FrmSearchUnits.show
            
End If
End Sub

Private Sub DcbUnitDDis_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Item = 4
Load FrmSearchUnits
            FrmSearchUnits.show
            
End If
End Sub

Private Sub dcbUnitDis_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Item = 5
Load FrmSearchUnits
            FrmSearchUnits.show
            
End If
End Sub

Private Sub DcbUnitDit_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Item = 3
Load FrmSearchUnits
            FrmSearchUnits.show
            
End If
End Sub

Private Sub DcbUnitG_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Item = 6
Load FrmSearchUnits
            FrmSearchUnits.show
            
End If
End Sub

Private Sub DcbUnitGroup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Item = 1
Load FrmSearchUnits
            FrmSearchUnits.show
            
End If
End Sub

Private Sub DcItem1_Change()
DcItem1_Click (0)
End Sub

Private Sub DcItem1_Click(Area As Integer)
      If val(DcItem1.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 Dim ownerid As Integer

    GetItemIDFromCode , val(DcItem1.BoundText), 1, EmpCode
    
    Me.Text1.text = EmpCode
End Sub

Private Sub DcItem1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Item = 5
Load FrmItemSearch
          
            FrmItemSearch.show
            
End If
End Sub
Sub Retrivetitems()
    Dim i    As Integer
    Dim j    As Integer
    Dim k    As Integer
 
    Dim Msg  As String
    Dim bool As Boolean
    Dim Rs1  As ADODB.Recordset
    Dim sql  As String
    bool = True
  
    '   FG.Rows = 10000
    FgItemPloice.Enabled = True
 
    With FgItemPloice

        If XPOptShowType(2).value = True Then
            If val(DcbItemDit.BoundText) = 0 Then
                MsgBox "Ì—ÃÏ «Œ Ì«— «·’‰›"
                DcbItemDit.SetFocus
                Exit Sub
            End If

            j = .rows
            .rows = .rows + 1
            Set Rs1 = New ADODB.Recordset

            For i = j To .rows - 1
                '  .TextMatrix(i, .ColIndex("name")) = DcItem1.text
                '.TextMatrix(i, .ColIndex("id")) = DcItem1.BoundText
                '  .TextMatrix(i, .ColIndex("Ser")) = i
                ' .TextMatrix(i, .ColIndex("uniteId")) = DcbUnitG.BoundText
                '  .TextMatrix(i, .ColIndex("unite")) = DcbUnitG.text
                sql = "  SELECT      dbo.TblItems.ItemID, dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee"
                sql = sql & "   FROM         dbo.Groups RIGHT OUTER JOIN"
                sql = sql & "                  dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID"
                sql = sql & "  Where (dbo.TblItems.ItemID =" & val(Me.DcbItemDit.BoundText) & ")"
                Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Rs1.RecordCount > 0 Then
                    .TextMatrix(i, .ColIndex("GropId")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
          
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
   
                    Else
                        .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
                    End If
                End If

                .TextMatrix(i, .ColIndex("id")) = Me.DcbItemDit.BoundText
                .TextMatrix(i, .ColIndex("name")) = Me.DcbItemDit.text
                .TextMatrix(i, .ColIndex("fullcode")) = Me.TxtCode.text
           
                .TextMatrix(i, .ColIndex("unite")) = Me.DcbUnitDit.text
                .TextMatrix(i, .ColIndex("uniteid")) = Me.DcbUnitDit.BoundText
                .TextMatrix(i, .ColIndex("price")) = Me.TxtPriceDit.text
                .TextMatrix(i, .ColIndex("amount")) = Me.TxtAmountDit.text

                If Me.DcbTypePoliceyDit.ListIndex = 2 Then
                    DcbItemDDis.BoundText = DcbItemDit.BoundText
                End If

                '.TextMatrix(i, .ColIndex("typedisid")) = Me.DcbTypePoliceyDit.ListIndex
                .TextMatrix(i, .ColIndex("typedis")) = Me.DcbTypePoliceyDit.ListIndex
                .TextMatrix(i, .ColIndex("discount")) = Me.TxtRateD.text
                ' .TextMatrix(i, .ColIndex("itemdisid")) = Me.DcbItemDDis.BoundText
                '  .TextMatrix(i, .ColIndex("itemdis")) = Me.DcbItemDDis.text
                '.TextMatrix(i, .ColIndex("unitdis")) = Me.DcbUnitDDis.text
                '.TextMatrix(i, .ColIndex("unitdisid")) = Me.DcbUnitDDis.BoundText
                .TextMatrix(i, .ColIndex("amountdis")) = Me.TxtAmountDDis.text
                .TextMatrix(i, .ColIndex("pricedis")) = Me.TxtPriceDDis.text
        
            Next i

            DcbItemDit.text = ""
            DcbUnitDit.text = ""
            TxtAmountDit.text = ""
            TxtPriceDit.text = ""
            DcbTypePoliceyDit.text = ""
            TxtRateD.text = ""
        End If
       
        If XPOptShowType(0).value = True Then
            If RdAllPolice.value = False Then
                MsgBox "ÌÃ»  ÕœÌœ «·”Ì«”Â"
                Exit Sub
            End If

            Set Rs1 = New ADODB.Recordset
            sql = " SELECT     dbo.TblItems.Fullcode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, "
            sql = sql & "           dbo.TblUnites.UnitNamee , dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode, dbo.Groups.GroupNamee"
            sql = sql & "            FROM         dbo.Groups RIGHT OUTER JOIN"
            sql = sql & "            dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID LEFT OUTER JOIN"
            sql = sql & "            dbo.TblUnites RIGHT OUTER JOIN"
            sql = sql & "            dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
            Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Rs1.RecordCount > 0 Then

                j = .rows
                .rows = .rows + Rs1.RecordCount

                For i = j To .rows - 1
            
                    .TextMatrix(i, .ColIndex("typedis")) = Me.DcbtypPolicep.ListIndex
                    .TextMatrix(i, .ColIndex("discount")) = Me.txtRate.text
                    .TextMatrix(i, .ColIndex("amountdis")) = Me.TxtAmountDis.text
                    .TextMatrix(i, .ColIndex("pricedis")) = Me.TxtPriceDis.text

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
            
                        .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
                        .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                        .TextMatrix(i, .ColIndex("unite")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
                    Else
                        .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
                        .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                        .TextMatrix(i, .ColIndex("unite")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
                    End If

                    .TextMatrix(i, .ColIndex("GropId")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs1("ItemID").value), "", Rs1("ItemID").value)
                    .TextMatrix(i, .ColIndex("Ser")) = i
      
                    .TextMatrix(i, .ColIndex("uniteId")) = IIf(IsNull(Rs1("UnitID").value), "", Rs1("UnitID").value)
                    Rs1.MoveNext
        
                Next i

            End If
       
        End If

        Dim GROUPIDS As String
        
        If XPOptShowType(1).value = True Then
            If RdAllPolice.value = False Then
                MsgBox "ÌÃ»  ÕœÌœ «·”Ì«”Â"
                Exit Sub
            End If

            For k = 1 To ListGroupSelected.ListCount

                Set Rs1 = New ADODB.Recordset
                '   sql = " SELECT * from  TblItems where GroupID =" & ListGroupSelected.ItemData(k - 1) & ""
                GROUPIDS = GetallChilddata(ListGroupSelected.ItemData(k - 1))

                If mId(GROUPIDS, 1, 1) = "," And Len(GROUPIDS) > 2 Then
                    GROUPIDS = mId(GROUPIDS, 2, Len(GROUPIDS))
                End If
   
                '       If Len(GROUPIDS) > 2 Then GROUPIDS = Mid(GROUPIDS, 2, Len(GROUPIDS))
                Debug.Print GROUPIDS

                If GROUPIDS = "" Then GROUPIDS = ListGroupSelected.ItemData(k - 1)
                sql = " SELECT dbo.TblItems.Fullcode,     dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.ItemID, dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, "
                sql = sql & "           dbo.TblUnites.UnitNamee , dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode, dbo.Groups.GroupNamee"
                sql = sql & "            FROM         dbo.Groups RIGHT OUTER JOIN"
                sql = sql & "            dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID LEFT OUTER JOIN"
                sql = sql & "            dbo.TblUnites RIGHT OUTER JOIN"
                sql = sql & "            dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID"
        
                sql = sql & "  where dbo.TblItems.GroupID IN (" & ListGroupSelected.ItemData(k - 1) & "," & GROUPIDS & ")"
                sql = sql & " order by TblItems.GroupID"
                '(GetallChilddata
                Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Rs1.RecordCount > 0 Then
                    j = .rows
                    .rows = .rows + Rs1.RecordCount

                    For i = j To .rows - 1
                        .TextMatrix(i, .ColIndex("typedis")) = val(Me.DcbtypPolicep.ListIndex)
                        .TextMatrix(i, .ColIndex("discount")) = Me.txtRate.text
                        .TextMatrix(i, .ColIndex("amountdis")) = Me.TxtAmountDis.text
                        .TextMatrix(i, .ColIndex("pricedis")) = Me.TxtPriceDis.text
                        .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)

                        If SystemOptions.UserInterface = ArabicInterface Then
                            .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
                            .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                            .TextMatrix(i, .ColIndex("unite")) = IIf(IsNull(Rs1("UnitName").value), "", Rs1("UnitName").value)
                        Else
                            .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
                            .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
                            .TextMatrix(i, .ColIndex("unite")) = IIf(IsNull(Rs1("UnitNamee").value), "", Rs1("UnitNamee").value)
                        End If

                        .TextMatrix(i, .ColIndex("GropId")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
                        .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs1("ItemID").value), "", Rs1("ItemID").value)
                        .TextMatrix(i, .ColIndex("Ser")) = i
      
                        .TextMatrix(i, .ColIndex("uniteId")) = IIf(IsNull(Rs1("UnitID").value), "", Rs1("UnitID").value)
                        Rs1.MoveNext
    
                    Next i

                End If
       
            Next k

        End If

    End With
    '***********************
    Dim Row As Integer
    For Row = 1 To FgItemPloice.rows - 1
        Dim mItemId As Long
        mItemId = val(FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("id")))
        If mItemId > 0 Then
            Dim ItemData As ItemsOverData
            ItemData = GetItemOverData(mItemId)
            FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colResultValue")) = ItemData.FinalQty
            FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colTotalQtyP")) = ItemData.QtyPeriod
            FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colAvgQtyD")) = ItemData.QtyDay
            On Error GoTo hErr
            txtResultValue = FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colResultValue"))
            txtTotalQtyP = FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colTotalQtyP"))
              txtAvgQtyD = FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colAvgQtyD"))
hErr:
            'ignor
        End If
    Next
        
    '************************
    ReLineGrid
End Sub

Private Sub ENDDATE_Change()
 On Error Resume Next

    If IsDate(StartDate) And IsDate(EndDate) Then
        txtPeriod2 = DateDiff("d", StartDate, EndDate)
    Else
        txtPeriod2 = "0"
    End If
End Sub

Private Sub FgItemPloice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With FgItemPloice
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If
        ReLineGrid
 
    End With
    If Col = FgItemPloice.ColIndex("Fullcode") Or _
       Col = FgItemPloice.ColIndex("name") Then
        Dim mItemId As Long
        mItemId = val(FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("id")))
        If mItemId > 0 Then
            Dim ItemData As ItemsOverData
            ItemData = GetItemOverData(mItemId)
            FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colResultValue")) = ItemData.FinalQty
            FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colTotalQtyP")) = ItemData.QtyPeriod
            FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colAvgQtyD")) = ItemData.QtyDay
            On Error GoTo hErr
            txtResultValue = FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colResultValue"))
            txtTotalQtyP = FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colTotalQtyP"))
            txtAvgQtyD = FgItemPloice.TextMatrix(Row, FgItemPloice.ColIndex("colAvgQtyD"))
hErr:
        End If
    End If
     
End Sub

Private Sub FgItemPloice_AfterRowColChange(ByVal OldRow As Long, _
                                           ByVal OldCol As Long, _
                                           ByVal NewRow As Long, _
                                           ByVal NewCol As Long)
    On Error Resume Next
    txtResultValue = FgItemPloice.TextMatrix(NewRow, FgItemPloice.ColIndex("colResultValue"))
    txtTotalQtyP = FgItemPloice.TextMatrix(NewRow, FgItemPloice.ColIndex("colTotalQtyP"))
    txtAvgQtyD = FgItemPloice.TextMatrix(NewRow, FgItemPloice.ColIndex("colAvgQtyD"))
End Sub

Private Sub FgItemPloice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)


    'On Error GoTo ErrTrap

    With Me.FgItemPloice

        Select Case .ColKey(Col)

                 Case "itemdis"
                  LngRow = Row

 LngCol = Col
             ' ItemProductionDate Row, Col, , 1
                Load FrmItemShowDet
                FrmItemShowDet.show vbModal

                    
                End Select
                End With
End Sub

Private Sub FgItemPloice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With Me.FgItemPloice

        Select Case .ColKey(Col)

                 Case "itemdis"
    
            .ColComboList(.ColIndex("itemdis")) = "..."
            End Select
           End With
End Sub

Private Sub ISButton1_Click()

Retrivetitems

End Sub

Private Sub Label8_Click()
Dim GROUPIDS, sql As String
Dim Rs1  As ADODB.Recordset
Dim i, k As Integer
 If Me.XPOptShowType(1).value = True Then
 If ListGroupAll.ListIndex > -1 Then
    ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
             
    ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)

            End If
            End If
End Sub
Private Sub Label6_Click()
    ListGroupSelected.Clear
End Sub
Private Sub Label5_Click()

    If ListGroupSelected.ListIndex > -1 Then
        ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
    End If

End Sub
Private Sub Label7_Click()
    Dim i As Integer
    If Me.XPOptShowType(1).value = True Then
    ListGroupSelected.Clear

    For i = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(i)
        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
    Next i
End If
End Sub
'Private Sub RdAllPolice_Click()
'saIf Me.RdAllPolice.value = True Then
'sa Me.Fra(7).Enabled = False
'sa Me.Fra(0).Enabled = True
'sa Me.Fra(3).Enabled = True
'sa Else
'sa Me.Fra(7).Enabled = True
'sa .Fra(0).Enabled = False
'saMe.Fra(3).Enabled = False
'saEnd If
'End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).value = True Then
    Frame3.Visible = False
    Frame2.Visible = False
    EleHeader.Caption = "⁄—Ê÷ «·«’‰«›"
ElseIf Option1(1).value = True Then
    Frame3.Visible = True
    Frame2.Visible = True
    EleHeader.Caption = "«⁄œ«œ«  ﬂ„Ì«  «Ê«„— «·‘—«¡"
End If
End Sub

Private Sub STARTDATE_Change()
 On Error Resume Next

    If IsDate(StartDate) And IsDate(EndDate) Then
        txtPeriod2 = DateDiff("d", StartDate, EndDate)
    Else
        txtPeriod2 = "0"
    End If
   
End Sub

Private Sub Startdate2_Change()
     
callcPer
End Sub

Private Sub enddate2_Change()
    callcPer

End Sub
Sub callcPer()

  On Error Resume Next

    If IsDate(Startdate2) And IsDate(enddate2) Then
        txtPeriod = DateDiff("d", Startdate2, enddate2)
    Else
        txtPeriod = "0"
    End If
    txtTotalPer = CStr(val(txtPeriod) + val(txtProdArrive))
End Sub
'Private Sub RdPrivatePolice_Click()
'If Me.RdPrivatePolice.value = True Then
'XPOptShowType(0).value = False
'XPOptShowType(1).value = False
'XPOptShowType(2).value = False
'Me.Fra(7).Enabled = True
'Me.Fra(0).Enabled = False
'Me.Fra(3).Enabled = False
'End If
'End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
   Dim ItemID As Integer

    If KeyAscii = vbKeyReturn Then
        GetItemIDFromCode Me.Text1.text, ItemID
        DcItem1.BoundText = ItemID
    End If
End Sub

'Private Sub TxtAmountDit_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF3 Then
'ch = False
'Load FrmSearchItemShow
'            FrmSearchItemShow.show
            
'End If
'End Sub

'Private Sub TxtNameShow_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF3 Then
'ch = False
'Load FrmSearchItemShow
'            FrmSearchItemShow.show
'
'End If
'End Sub

'Private Sub TxtPriceDit_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF3 Then
'ch = False
'Load FrmSearchItemShow
'            FrmSearchItemShow.show
            
'End If
'End Sub

'    If KeyCode = vbKeyF3 Then
'        FrmEmployeeSearch.lbltype = 9
''        FrmEmployeeSearch.show
'
'   End If

'End Sub

'Private Sub DcboEmpName_Click(Area As Integer)
'   On Error Resume Next
''       If val(DcboEmpName.BoundText) = 0 Then Exit Sub
'
'    Dim EmpCode  As String
 
''    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
'   TxtSearchCode.text = EmpCode
'
'  If Me.TxtModFlg = "R" Then Exit Sub
   
''
'  Dim StrSQL As String
'
'
        
'        Dim issuedate As Date
'        Dim depid As Double
'        Dim specid As Double
''        Dim JobTypeID As Double
'       Dim gradeID As Double
'       Dim Account_code2 As String
'          Dim Account_Code  As String
'       Dim Balance As String
'       Dim endContractPerMonth As Double
'       Dim national As String
'       Dim project As Integer
'      Dim pasid As String
''     Dim iqamaid As String
'    Dim placeiqama As String
'    Dim endiq As String
'      get_employee_information val(Me.DcboEmpName.BoundText), issuedate, depid, specid, JobTypeID, gradeID, Account_code2, Account_Code, endContractPerMonth, national, , , project, pasid, iqamaid, placeiqama, , endiq
        
'    WriteCustomerBalPublic Account_code2, Balance
          
'lbl(22).Caption = val(Balance)

'       WriteCustomerBalPublic Account_Code, Balance
          
' lbl(21).Caption = val(Balance)
' l'bl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
'   DBIssueDate.value = issuedate
' DcboEmpDepartments.BoundText = project
'   DcboSpecifications.BoundText = gradeID
'   Me.TxtIqFrom.text = placeiqama
'   DcbEmpNation.text = national
'      DcboJobsType.BoundText = JobTypeID
'      TxtIqama.text = iqamaid
'      Me.XpDtEnd.value = endiq
'     TxtPas.text = pasid
        
'   lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
        
'End If

'End Sub

'Private Sub Command1_Click()
'  Dim i As Integer
'  Dim j As Integer
'  Dim k As Integer
 
'  Dim Msg As String
''  Dim bool As Boolean
' Dim rs1 As ADODB.Recordset
' Dim sql As String
' bool = True
'
'     If ListStoreSelected.ListCount = 0 Then
'      If SystemOptions.UserInterface = ArabicInterface Then
'           Msg = "Õœœ     „Œ“‰ Ê«Õœ ⁄·Ï «·«ﬁ· " & Chr(13)
'    Else
'    Msg = "Select At Least One Store " & Chr(13)
'    End If
'           MsgBox Msg, vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'
'           SendKeys "{F4}"
'           Exit Sub
'       End If
'       fg.Rows = 10000
'       fg.Enabled = True
'Set rs1 = New ADODB.Recordset
'  sql = " SELECT * from  TblItems "
' rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'If (XPOptShowType(0).value = True Or XPOptShowType(1).value = True) And bool = True Then
'   bool = False
'             Fg.Rows = (ListStoreSelected.ListCount) * rs1.RecordCount
               
'Fg.Enabled = True
'Else
'If (XPOptShowType(0).value = True Or XPOptShowType(1).value = True) And (bool = False) Then
'                  Fg.Rows = Fg.Rows + ((ListStoreSelected.ListCount) * rs1.RecordCount)
                
'Fg.Enabled = True
'End If
'End If
'   Fg.Rows = Fg.Rows + 1

'If (XPOptShowType(2).value = True) And fg.Rows < 2 Then
'
'      Else
'          fg.Rows = ListStoreSelected.ListCount + 1
'      fg.Enabled = True
'       End If
 
'   For i = 1 To ListStoreSelected.ListCount
'   If XPOptShowType(2).value = True Then
''          coun = coun + 1
'     fg.TextMatrix(count, fg.ColIndex("serial")) = coun
''      fg.TextMatrix(coun, fg.ColIndex("StoreName")) = ListStoreSelected.List(i - 1)
'     fg.TextMatrix(coun, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i - 1)
'              fg.TextMatrix(coun, fg.ColIndex("ItemName")) = Me.DcItem1.text
'    fg.TextMatrix(coun, fg.ColIndex("ItemID")) = Me.DcItem1.BoundText
'     End If
       
'        If XPOptShowType(0).value = True Then

' Set rs1 = New ADODB.Recordset
'        sql = " SELECT * from  TblItems "
'        rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
''
'   If rs1.RecordCount > 0 Then

'       For j = 1 To rs1.RecordCount
'coun = coun + 1
'            If SystemOptions.UserInterface = ArabicInterface Then
           
'              fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(rs1("ItemName").value), "", rs1("ItemName").value)
'            Else
'                fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(rs1("ItemNamee").value), "", rs1("ItemNamee").value)
'            End If

'         fg.TextMatrix(coun, fg.ColIndex("ItemID")) = rs1("ItemID").value
                 
'       fg.TextMatrix(coun, fg.ColIndex("serial")) = coun
'        fg.TextMatrix(coun, fg.ColIndex("StoreName")) = ListStoreSelected.List(i - 1)
'        fg.TextMatrix(coun, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i - 1)
'        rs1.MoveNext
'        Next j

'    End If
       
'        End If
'          If XPOptShowType(1).value = True Then
'          For k = 1 To ListGroupSelected.ListCount

'    Set rs1 = New ADODB.Recordset
'           sql = " SELECT * from  TblItems where GroupID =" & ListGroupSelected.ItemData(k - 1) & ""
'           rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

'    If rs1.RecordCount > 0 Then

'        For j = 1 To rs1.RecordCount
'coun = coun + 1
'            If SystemOptions.UserInterface = ArabicInterface Then
           
'              fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(rs1("ItemName").value), "", rs1("ItemName").value)
'            Else
'                fg.TextMatrix(coun, fg.ColIndex("ItemName")) = IIf(IsNull(rs1("ItemNamee").value), "", rs1("ItemNamee").value)
''            End If
'
'         fg.TextMatrix(coun, fg.ColIndex("ItemID")) = rs1("ItemID").value
''            fg.TextMatrix(coun, fg.ColIndex("GroupID")) = rs1("GroupID").value
'             fg.TextMatrix(coun, fg.ColIndex("GroupName")) = ListGroupSelected.List(k - 1)
'      fg.TextMatrix(coun, fg.ColIndex("serial")) = coun
'       fg.TextMatrix(coun, fg.ColIndex("StoreName")) = ListStoreSelected.List(i - 1)
'       fg.TextMatrix(coun, fg.ColIndex("StoreID")) = ListStoreSelected.ItemData(i - 1)
''       rs1.MoveNext
'      Next j
'
'    End If
       
'         Next k
'        End If
'    Next i
'    If XPOptShowType(0).value = True Or XPOptShowType(1).value = True Then
'    fg.Rows = coun + 1
'    End If
'    ReLineGrid
'End Sub

'Private Sub Label2_Click()
'    Dim i As Integer
'    ListStoreSelected.Clear
'''
' '   For i = 0 To ListStoreall.ListCount - 1
'       ListStoreSelected.AddItem ListStoreall.List(i)
'       ListStoreSelected.ItemData(i) = ListStoreall.ItemData(i)
'   Next i
'
'End Sub
'Private Sub Label5_Click()
'
''    If ListGroupSelected.ListIndex > -1 Then
'       ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
'   End If
''
'End Sub
'Private Sub Label6_Click()
'    ListGroupSelected.Clear
'End Sub
'Private Sub Label7_Click()
'    Dim i As Integer
'    If Me.XPOptShowType(1).value = True Then
''    ListGroupSelected.Clear
'
'    For i = 0 To ListGroupAll.ListCount - 1
'        ListGroupSelected.AddItem ListGroupAll.List(i)
'        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
'    Next i
'End If
'End Sub
'Private Sub Label8_Click()
'If Me.XPOptShowType(1).value = True Then
'' If ListGroupAll.ListIndex > -1 Then
''   ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
             
'  ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)
'          End If
'          End If
'End Sub
'Private Sub Label4_Click()
'
'    If ListStoreSelected.ListIndex > -1 Then
    
'        ListStoreSelected.RemoveItem ListStoreSelected.ListIndex
'    End If

'End Sub
'Private Sub Label3_Click()
'    ListStoreSelected.Clear
'End Sub

'Private Sub LblSelect_Click()
'If ListStoreall.ListIndex > -1 Then
'    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
'
'    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
'            End If
'End Sub

'Private Sub ListGroupAll_Click()
' If XPOptShowType(1).value = True Then
'        Frame11.Enabled = True
'    Else
'        Frame11.Enabled = False
'    End If
'End Sub

'Private Sub XPDtbTrans_Change()
'
'    If Trim(TxtNoteSerial1.text) <> "" Then
'        oldtxtNoteSerial1.text = TxtNoteSerial1.text
'    End If
'
'    TxtNoteSerial.text = ""
'    TxtNoteSerial1.text = ""

'End Sub

'Private Sub dcBranch_Click(Area As Integer)
 
' TxtNoteSerial.text = ""
' TxtNoteSerial1.text = ""
'End Sub

Private Sub Form_Load()
    FrmItemSearch.RetrunType = 17
    Dim Dcombos As ClsDataCombos
    'Set Dcombos = New ClsDataCombos
    Dim StrSQL  As String
    Dim GrdBack As ClsBackGroundPic

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        FgItemPloice.ColComboList(FgItemPloice.ColIndex("typedis")) = "#0; Œ’„ ﬁÌ„Â|#1; Œ’„ ‰”»Â|#2; Œ’„ ﬂ„ÌÂ „‰ ‰›” «·’‰›|#3; Œ’„ ﬂ„ÌÂ „‰ ’‰› «Œ—"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        FgItemPloice.ColComboList(FgItemPloice.ColIndex("typedis")) = "#0;Dis Value |#1;Dis Rate |#2; Dis Same Item|#3; Dis Another Item"
    End If
            
    FillMylist

    If SystemOptions.UserInterface = EnglishInterface Then
        DcbtypPolicep.AddItem "Dis Value"
        DcbtypPolicep.AddItem "Dis Rate"
        DcbtypPolicep.AddItem "Dis Same Item"
        DcbtypPolicep.AddItem "Dis Another Item"
        DcbtypPolicep.AddItem "Special Offers"
     
        DcbTypePoliceyDit.AddItem "Dis Value"
        DcbTypePoliceyDit.AddItem "Dis Rate"
        DcbTypePoliceyDit.AddItem "Dis Same Item"
        DcbTypePoliceyDit.AddItem "Dis Another Item"
    Else
   
        Me.DcbtypPolicep.AddItem "Œ’„ ﬁÌ„…"
        Me.DcbtypPolicep.AddItem "Œ’„ ‰”»…"
        Me.DcbtypPolicep.AddItem "Œ’„ ﬂ„Ì…„‰ ‰›” «·’‰›"
        Me.DcbtypPolicep.AddItem "Œ’„ ﬂ„Ì… „‰ ’‰› «Œ—"
        Me.DcbtypPolicep.AddItem "⁄—÷ Œ«’"
  
        Me.DcbTypePoliceyDit.AddItem "Œ’„ ﬁÌ„…"
        Me.DcbTypePoliceyDit.AddItem "Œ’„ ‰”»…"
        Me.DcbTypePoliceyDit.AddItem "Œ’„ ﬂ„Ì…„‰ ‰›” «·’‰›"
        Me.DcbTypePoliceyDit.AddItem "Œ’„ ﬂ„ÌÂ „‰ ’‰› «Œ—"
    End If
 
    With CboFromPrice
        .Clear
        .AddItem "«ﬁ· ”⁄—"
        .AddItem "«⁄·Ì ”⁄—"
    End With
 
    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
    '  Frame10.Enabled = False
    ' Frame11.Enabled = False
    'Me.XPOptShowType(0).value = True
    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
     
    Resize_Form Me
    AddTip
    
    ' GetItemsUnits DcbItem
    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.DcbBranch1
     Dcombos.GetPOS Me.dcPOS
    Dcombos.GetBranches Me.DcbBranch
    Dcombos.GetItemSGroupsupdate DcbGroup, True
    Dcombos.GetItemsNamesupdate Me.DcbItem
    Dcombos.GetItemsNamesupdate Me.DcbItemBisc1
    ' Dcombos.GetItemsNames Me.DcbItemBisc2
   
    Dcombos.GetItemsNamesupdate Me.DcbItemDDis
    Dcombos.GetItemsNamesupdate Me.DcbItemDis
    Dcombos.GetItemsNamesupdate Me.DcbItemDit
    Dcombos.GetItemsNamesupdate Me.DcItem1
    
    Dcombos.GetItemsUnits Me.dcbUnitDis
    Dcombos.GetItemsUnits Me.DcbUnit
    Dcombos.GetItemsUnits Me.dcbUnitBisc1
    Dcombos.GetItemsUnits Me.DcbUnitG
    Dcombos.GetItemsUnits Me.DcbUnitDDis
    Dcombos.GetItemsUnits Me.DcbUnitGroup
    Dcombos.GetItemsUnits Me.DcbUnitDit

    Dcombos.GetCustomersSuppliers 3, Me.DBCboClientName, True

    If SystemOptions.usertype <> UserAdminAll Then
        Me.DcbBranch.Enabled = False
    End If

    SetDtpickerDate Me.XPDtbTrans
    'YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblItemShows  where IsNull(TransType,0)=" & mIndex & "   Order By id"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
    Me.TxtModFlg.text = "R"
    Retrive

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    If mIndex = 0 Then
        Option1(0).value = True
    ElseIf mIndex = 1 Then
        Option1(1).value = True
    End If

    Option1_Click 1
    Exit Sub

ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    '    Label1.Visible = False

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    Cmd(21).Caption = "Delete"
    Cmd(13).Caption = "Delete"
    Cmd(12).Caption = "Delete"
    Cmd(14).Caption = "Delete All"
    Cmd(15).Caption = "Delete All"
    Cmd(16).Caption = "Delete All"
    Me.Caption = "Screen Shows Items  "
    lbl(9).Caption = " Name Show"
    EleHeader.Caption = Me.Caption
    lbl(3).Caption = "No Show"
    lbl(1).Caption = "Date"
    lbl(4).Caption = "Branch"
    lbl(17).Caption = "Branch"
    lbl(2).Caption = "Start "
    lbl(5).Caption = "End "
    Fra(0).Caption = "Data items or groups of items, which is applied to the display"
    Fra(1).Caption = "Data of Branch, which is applied to the display"
    Fra(3).Caption = "Discount Policy"
    lbl(15).Caption = "Group Name "
    lbl(10).Caption = "Item Name "
    lbl(11).Caption = "Item Name "
    lbl(14).Caption = "Unite "
    lbl(13).Caption = "Unite "
    lbl(16).Caption = "Unite "
    '        lbl(18).Caption = "Amount "
    Me.ChAllBranch.Caption = "All Branch"
    BtonAdd.Caption = "Insert"
    BtonAdd1.Caption = "Insert"
    BtonAdd2.Caption = "Insert"
    BtonAdd3.Caption = "Insert"
    Me.RdAllPolice.Caption = "Per the policy of total items specified"
    Me.RdPrivatePolice.Caption = "Customize the display according to each class"
    RdAllPolice.RightToLeft = False
    RdPrivatePolice.RightToLeft = False
           
    lbl(28).Caption = "TypeDiscount "
    lbl(47).Caption = "TypeDiscount "
    lbl(51).Caption = "Item Name "
    lbl(22).Caption = "Item Discount "
    lbl(42).Caption = "Item Discount "
    lbl(19).Caption = "Unit "
    lbl(50).Caption = "Unit "
    lbl(39).Caption = "Unit "
    lbl(20).Caption = "Amount "
    lbl(24).Caption = "Amount "
    lbl(40).Caption = "Amount "
    lbl(49).Caption = "Amount "
    lbl(26).Caption = "Price "
    lbl(21).Caption = "Price "
    lbl(45).Caption = "Price "
    lbl(48).Caption = "Price "
    XPTab301.Caption = "Data Show Items"

    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
    With Me.FgBranch
        .TextMatrix(0, .ColIndex("id")) = "Branch Code"
        .TextMatrix(0, .ColIndex("name")) = "Branch Name"
    End With
      With Me.grdPos
        .TextMatrix(0, .ColIndex("id")) = "POS Code"
        .TextMatrix(0, .ColIndex("name")) = "POS Name"
    End With

    With Me.FgItems
        .TextMatrix(0, .ColIndex("Ser")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "Item No"
        .TextMatrix(0, .ColIndex("name")) = "Item Name"
        .TextMatrix(0, .ColIndex("unite")) = " Unite"
        .TextMatrix(0, .ColIndex("group")) = "Group Name"
        .TextMatrix(0, .ColIndex("unite1")) = " Unite"
    End With
    With Me.FgItems
        .TextMatrix(0, .ColIndex("Ser")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "Item No"
        .TextMatrix(0, .ColIndex("name")) = "Item Name"
        .TextMatrix(0, .ColIndex("unite")) = " Unite"
        .TextMatrix(0, .ColIndex("group")) = "Group Name"
    End With
    With Me.FgItemPloice
        .TextMatrix(0, .ColIndex("Ser")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "Item No"
        .TextMatrix(0, .ColIndex("name")) = "Item Name"
        .TextMatrix(0, .ColIndex("unite")) = " Unite"
        .TextMatrix(0, .ColIndex("amount")) = "Amount"
        .TextMatrix(0, .ColIndex("typedis")) = "Type Discount"
        .TextMatrix(0, .ColIndex("discount")) = "Discount"
        .TextMatrix(0, .ColIndex("group")) = "GroupName"
        .TextMatrix(0, .ColIndex("untgroup")) = " Unite"
        .TextMatrix(0, .ColIndex("price")) = "Price"
        .TextMatrix(0, .ColIndex("itemdis")) = "ItemDis"
        .TextMatrix(0, .ColIndex("unitdis")) = "Unite"

        .TextMatrix(0, .ColIndex("amountdis")) = "Amount"
        .TextMatrix(0, .ColIndex("pricedis")) = "price"
    End With

End Sub



Private Sub Form_Paint()
    TTD.Destroy
End Sub

Private Sub Form_Resize()
    TTD.Destroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

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



Private Sub TXTCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If TxtCode.text = "" Then
            Me.DcbItemDit.BoundText = ""
        Else
            Me.DcbItemDit.BoundText = GetItemID(Trim$(Me.TxtCode.text))
        End If
    End If
    
End Sub

Private Sub txtMin_Change()
txtPlus = ""
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰"
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
          '  TxtAdvanceValue.Locked = True
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
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
        '    TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰(  ⁄œÌ· )"
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
        '    TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub



Private Sub txtPlus_Change()
txtMin = ""
End Sub

Private Sub txtProdArrive_Change()
callcPer
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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails  As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim RsDetails2 As ADODB.Recordset
    
    Dim i          As Integer
    Dim StrSQL     As String
    ListGroupSelected.Clear
    FgBranch.Clear flexClearScrollable, flexClearEverything
    
    FgBranch.rows = 1
    Me.FgBranch.Enabled = True
    
    grdPos.Clear flexClearScrollable, flexClearEverything
    grdPos.rows = 1
    Me.grdPos.Enabled = True
            
    Me.FgItemPloice.Clear flexClearScrollable, flexClearEverything
    FgItemPloice.rows = 1
    Me.FgItemPloice.Enabled = True
            
    '   Me.FgItems.Clear flexClearScrollable, flexClearEverything
    ' FgItems.Rows = 1
    ' Me.FgItems.Enabled = True
    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "id=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
    
    '''////////
    
    XPTxtID.text = rs("ID").value
    '///////////////////////new 20 11 2016
    Me.TxtSales.text = IIf(IsNull(rs("Sales").value), "", rs("Sales").value)
    Me.TxtGetFree.text = IIf(IsNull(rs("GetFree").value), "", rs("GetFree").value)
    Me.txtDiscount.text = IIf(IsNull(rs("Discount").value), "", rs("Discount").value)
               
    Me.CboFromPrice.ListIndex = IIf(IsNull(rs("FromPrice").value), -1, rs("FromPrice").value)
     
    '///////////////////////new 20 11 2016
                
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Me.DcbBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    Me.DcbUnitG.BoundText = IIf(IsNull(rs("UnitG").value), "", rs("UnitG").value)
    Me.TxtNameShow.text = IIf(IsNull(rs("NameShow").value), "", rs("NameShow").value)
    StartDate.value = IIf(IsNull(rs("StartSDate").value), Date, rs("StartSDate").value)
    EndDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    Me.DcItem1.BoundText = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
    Me.DcbGroup.BoundText = IIf(IsNull(rs("GroupID").value), "", rs("GroupID").value)
    Me.DcbUnit.BoundText = IIf(IsNull(rs("UnitItID").value), "", rs("UnitItID").value)
    Me.DcbItemDit.BoundText = IIf(IsNull(rs("ItemIDD").value), "", rs("ItemIDD").value)
    Me.DcbUnitDit.BoundText = IIf(IsNull(rs("UnitItIDD").value), "", rs("UnitItIDD").value)
    Me.TxtAmountDit.text = IIf(IsNull(rs("AmountD").value), "", rs("AmountD").value)
    Me.TxtPriceDit.text = IIf(IsNull(rs("PriceD").value), 0, rs("PriceD").value)
    '*******************************
    txtPeriod.text = val(rs!Period & "")
    txtPlus.text = val(rs!PLUS & "")
    txtMin.text = val(rs!Min & "")
    'txtPurQty.text = val(rs!purQty & "")
    txtAvgQtyD.text = val(rs!AvgQtyD & "")
    txtTotalQtyP.text = val(rs!TotalQtyP & "")
    txtResultValue.text = val(rs!ResultValue & "")
    '********************************
     
    DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)

    If rs("Sa").value = True Then
        opt_sa.value = 1
    ElseIf rs("Sa").value = False Or IsNull(rs("Sa").value) Then
        opt_sa.value = 0
    End If
         
    If rs("Su").value = True Then
        opt_su.value = 1
    ElseIf rs("Su").value = False Or IsNull(rs("Su").value) Then
        opt_su.value = 0
    End If
         
    If rs("mo").value = True Then
        opt_mo.value = 1
    ElseIf rs("mo").value = False Or IsNull(rs("mo").value) Then
        opt_mo.value = 0
    End If
         
    If rs("Tu").value = True Then
        opt_tu.value = 1
    ElseIf rs("Tu").value = False Or IsNull(rs("tu").value) Then
        opt_tu.value = 0
    End If
         
    If rs("We").value = True Then
        opt_We.value = 1
    ElseIf rs("We").value = False Or IsNull(rs("we").value) Then
        opt_We.value = 0
    End If
         
    If rs("Th").value = True Then
        opt_Th.value = 1
    ElseIf rs("Th").value = False Or IsNull(rs("th").value) Then
        opt_Th.value = 0
    End If
         
    If rs("Fr").value = True Then
        opt_Fr.value = 1
    ElseIf rs("Fr").value = False Or IsNull(rs("fr").value) Then
        opt_Fr.value = 0
    End If

    Me.DcbTypePoliceyDit.ListIndex = IIf(IsNull(rs("TypePoliceD").value), -1, rs("TypePoliceD").value)
    Me.DcbItemDDis.BoundText = IIf(IsNull(rs("ItemIDDDis").value), "", rs("ItemIDDDis").value)
    Me.DcbUnitDDis.BoundText = IIf(IsNull(rs("UnitItIDDDis").value), "", rs("UnitItIDDDis").value)
    Me.TxtAmountDDis.text = IIf(IsNull(rs("AmountDDis").value), "", rs("AmountDDis").value)
    Me.TxtPriceDDis.text = IIf(IsNull(rs("PriceDDis").value), 0, rs("PriceDDis").value)

    If rs("AllBranch").value = 0 Then
        Me.ChAllBranch.value = xtpUnchecked
    Else
        Me.ChAllBranch.value = xtpChecked
    End If

    DcbBranch1.BoundText = IIf(IsNull(rs("BranchID2").value), "", rs("BranchID2").value)

    If rs("AllPolice").value = 0 Then
        Me.RdAllPolice.value = False
    Else
        Me.RdAllPolice.value = True
    End If

    If rs("PrivatePolice").value = 0 Then
        Me.RdPrivatePolice.value = False
    Else
        Me.RdPrivatePolice.value = True
    End If

    If val(rs("Selected").value) = 1 Then
        XPOptShowType(0).value = True
    Else
        XPOptShowType(0).value = False
    End If

    If val(rs("Selected").value) = 2 Then
        XPOptShowType(1).value = True
    Else
        XPOptShowType(1).value = False
    End If

    If val(rs("Selected").value) = 3 Then
        XPOptShowType(2).value = True
    Else
        XPOptShowType(2).value = False
    End If

    Me.DcbtypPolicep.ListIndex = IIf(IsNull(rs("TypePoliceP").value), -1, rs("TypePoliceP").value)

    If Me.DcbtypPolicep.ListIndex = 4 Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
    End If
   
    Me.dcbUnitBisc1.BoundText = IIf(IsNull(rs("UnitBisc").value), "", rs("UnitBisc").value)
    Me.DcbItemBisc1.BoundText = IIf(IsNull(rs("ItemIDBisc").value), "", rs("ItemIDBisc").value)
    Me.TxtAmountBisc1.text = IIf(IsNull(rs("AmountBisc").value), "", rs("AmountBisc").value)
    Me.TxtAmountBisc2.text = IIf(IsNull(rs("AmountBisc").value), "", rs("AmountBisc").value)
    Me.TxtPriceBisc1.text = IIf(IsNull(rs("PriceBisc").value), "", rs("PriceBisc").value)
    Me.TxtPriceBisc2.text = IIf(IsNull(rs("PriceBisc").value), "", rs("PriceBisc").value)
    Me.TxtAmountDis.text = IIf(IsNull(rs("AmountDis").value), "", rs("AmountDis").value)
       
    Me.DcbItemDis.BoundText = IIf(IsNull(rs("ItemDis").value), "", rs("ItemDis").value)
    Me.dcbUnitDis.BoundText = IIf(IsNull(rs("UnitItDis").value), "", rs("UnitItDis").value)
  
    Me.TxtPriceDis.text = IIf(IsNull(rs("PriceDis").value), "", rs("PriceDis").value)
    Me.txtRate.text = IIf(IsNull(rs("Rate").value), "", rs("Rate").value)
    Me.TxtRateD.text = IIf(IsNull(rs("RateD").value), "", rs("RateD").value)
    Me.DcbUnitGroup.BoundText = IIf(IsNull(rs("UnitGroup").value), "", rs("UnitGroup").value)

    ''////////

    'Me.DcbOrderStatus.ListIndex = rs("LinkType").value
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    '   If IsNull(rs("posted").value) Then
    '                                                If SystemOptions.UserInterface = ArabicInterface Then
    '                                                Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
    '                                              Else
    '                                                Accredit.Caption = " send to Approval   "
    '                                            End If
    '                                            Accredit.Enabled = True
    'Else
    '                                               If SystemOptions.UserInterface = ArabicInterface Then
    '                                                  Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
    ''                                                Else
    '                                                Accredit.Caption = " sent to Approval   "
    '                                            End If
    '                                            Accredit.Enabled = False
    'End If
    
    Set RsDetails = New ADODB.Recordset
    StrSQL = " SELECT     dbo.TblItemShowDitailses.ID, dbo.TblItemShowDitailses.ID2, dbo.TblItemShowDitailses.uniteId, dbo.TblItemShowDitailses.unitdisid, dbo.TblItemShowDitailses.GropId,"
    StrSQL = StrSQL & " TblItemShowDitailses.Period ,"
    StrSQL = StrSQL & " TblItemShowDitailses.Plus  ,"
    StrSQL = StrSQL & " TblItemShowDitailses.Min ,"
    StrSQL = StrSQL & " TblItemShowDitailses.PurQty  ,"
    StrSQL = StrSQL & " TblItemShowDitailses.AvgQtyD ,"
    StrSQL = StrSQL & " TblItemShowDitailses.TotalQtyP  ,"
    StrSQL = StrSQL & " TblItemShowDitailses.ResultValue ,"
    StrSQL = StrSQL & "             dbo.TblItems.Fullcode,         dbo.Groups.GroupName, dbo.TblItemShowDitailses.typedisid, dbo.TblItemShowDitailses.amount, dbo.TblItemShowDitailses.amountdis,"
    StrSQL = StrSQL & "                      dbo.TblItemShowDitailses.discount, dbo.TblItemShowDitailses.pricedis, dbo.TblItemShowDitailses.Price, dbo.TblItemShowDitailses.Type,"
    StrSQL = StrSQL & "                      dbo.TblItemShowDitailses.ItemID, dbo.TblItemShowDitailses.ItemDisID, dbo.TblItemShowDitailses.InfITemSho, dbo.TblItems.ItemName, dbo.TblItems.ItemCode,"
    StrSQL = StrSQL & "                      dbo.TblItems.ItemNamee , dbo.TblUnites.Unitname, dbo.TblUnites.UnitNamee"
    StrSQL = StrSQL & " FROM         dbo.TblItemShowDitailses LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblItems ON dbo.TblItemShowDitailses.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.Groups ON dbo.TblItemShowDitailses.GropId = dbo.Groups.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.TblItemShowDitailses.uniteId = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & " where id2 = " & val(XPTxtID.text) & " and Type =0 And IsNull(TblItemShowDitailses.TransType,0) = " & mIndex

    'StrSQL = " select * from TblItemShowDitails where id2 = " & val(XPTxtID.text) & " and Type =0 "

    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst

        With Me.FgItemPloice
            .rows = .FixedRows + RsDetails.RecordCount

            For i = .FixedRows To .rows - 1
     
                .TextMatrix(i, .ColIndex("Ser")) = i
          
                .TextMatrix(i, .ColIndex("idd")) = IIf(IsNull(RsDetails("ID").value), "", RsDetails("ID").value)
                .TextMatrix(i, .ColIndex("itemssh")) = IIf(IsNull(RsDetails("InfITemSho").value), "", RsDetails("InfITemSho").value)
      
                .TextMatrix(i, .ColIndex("uniteId")) = IIf(IsNull(RsDetails("uniteId").value), "", RsDetails("uniteId").value)
                .TextMatrix(i, .ColIndex("GropId")) = IIf(IsNull(RsDetails("GropId").value), "", RsDetails("GropId").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDetails("ItemID").value), "", RsDetails("ItemID").value)
                .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsDetails("fullcode").value), "", RsDetails("fullcode").value)
            
                .TextMatrix(i, .ColIndex("typedis")) = IIf(IsNull(RsDetails("typedisid").value), "", RsDetails("typedisid").value)
         
                .TextMatrix(i, .ColIndex("unitdisid")) = IIf(IsNull(RsDetails("unitdisid").value), "", RsDetails("unitdisid").value)
                .TextMatrix(i, .ColIndex("amount")) = IIf(IsNull(RsDetails("amount").value), "", RsDetails("amount").value)
                .TextMatrix(i, .ColIndex("amountdis")) = IIf(IsNull(RsDetails("amountdis").value), "", RsDetails("amountdis").value)
                .TextMatrix(i, .ColIndex("discount")) = IIf(IsNull(RsDetails("discount").value), "", RsDetails("discount").value)
                .TextMatrix(i, .ColIndex("pricedis")) = IIf(IsNull(RsDetails("pricedis").value), "", RsDetails("pricedis").value)
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDetails("Price").value), "", RsDetails("Price").value)
                .TextMatrix(i, .ColIndex("itemdisid")) = IIf(IsNull(RsDetails("ItemDisID").value), "", RsDetails("ItemDisID").value)
                .TextMatrix(i, .ColIndex("group")) = IIf(IsNull(RsDetails("GroupName").value), "", RsDetails("GroupName").value)
                '***********************************
                .TextMatrix(i, .ColIndex("colPeriod")) = val(RsDetails!Period & "")
                .TextMatrix(i, .ColIndex("colPlus")) = val(RsDetails!PLUS & "")
                .TextMatrix(i, .ColIndex("colMin")) = val(RsDetails!Min & "")
                .TextMatrix(i, .ColIndex("colPurQty")) = val(RsDetails!purQty & "")
                .TextMatrix(i, .ColIndex("colAvgQtyD")) = val(RsDetails!AvgQtyD & "")
                .TextMatrix(i, .ColIndex("colTotalQtyP")) = val(RsDetails!TotalQtyP & "")
                .TextMatrix(i, .ColIndex("colResultValue")) = val(RsDetails!ResultValue & "")
                '***********************************
          
                If SystemOptions.UserInterface = EnglishInterface Then
                    .TextMatrix(i, .ColIndex("unite")) = IIf(IsNull(RsDetails("UnitNamee").value), "", RsDetails("UnitNamee").value)
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("ItemNamee").value), "", RsDetails("ItemNamee").value)
                    '   .TextMatrix(i, .ColIndex("itemdis")) = IIf(IsNull(RsDetails("ItemNameDese").value), "", RsDetails("ItemNameDese").value)
                    '.TextMatrix(i, .ColIndex("unitdis")) = IIf(IsNull(RsDetails("UnitNameDese").value), "", RsDetails("UnitNameDese").value)
                Else
                    .TextMatrix(i, .ColIndex("unite")) = IIf(IsNull(RsDetails("UnitName").value), "", RsDetails("UnitName").value)
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("ItemName").value), "", RsDetails("ItemName").value)
                    '   .TextMatrix(i, .ColIndex("itemdis")) = IIf(IsNull(RsDetails("ItemNameDes").value), "", RsDetails("ItemNameDes").value)
                    '.TextMatrix(i, .ColIndex("unitdis")) = IIf(IsNull(RsDetails("UnitNameDes").value), "", RsDetails("UnitNameDes").value)
         
                End If

                RsDetails.MoveNext
            Next i

        End With

    End If

    '''\\\\\\\\\\\\\\\\\\\

    Set RsDetails = New ADODB.Recordset
    StrSQL = " SELECT     dbo.TblItemShowDitailses.ID, dbo.TblItemShowDitailses.ID2, dbo.TblItemShowDitailses.Type, dbo.TblItemShowDitailses.BrnchID, dbo.TblBranchesData.branch_name,"
    StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_nameE , dbo.TblBranchesData.branch_Code"
    StrSQL = StrSQL & " FROM         dbo.TblItemShowDitailses LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblItemShowDitailses.BrnchID = dbo.TblBranchesData.branch_id"
    StrSQL = StrSQL & " where id2 = " & val(XPTxtID.text) & " and Type =1 And TblItemShowDitailses.TransType = " & mIndex

    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst

        With Me.FgBranch
            .rows = .FixedRows + RsDetails.RecordCount

            For i = .FixedRows To .rows - 1
     
                .TextMatrix(i, .ColIndex("Ser")) = i
          
                .TextMatrix(i, .ColIndex("branchid")) = IIf(IsNull(RsDetails("BrnchID").value), "", RsDetails("BrnchID").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDetails("branch_Code").value), "", RsDetails("branch_Code").value)

                If SystemOptions.UserInterface = EnglishInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("branch_namee").value), "", RsDetails("branch_namee").value)
         
                Else
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
         
                End If
          
                RsDetails.MoveNext
            Next i

        End With

    End If

    '''
    Set RsDetails = Nothing
    '***********************
    
    '''\\\\\\\\\\\\\\\\\\\

    Set RsDetails = New ADODB.Recordset
    StrSQL = " SELECT     dbo.TblItemShowDitailses.ID, dbo.TblItemShowDitailses.ID2, "
    StrSQL = StrSQL & " dbo.TblItemShowDitailses.type , dbo.TblItemShowDitailses.brnchid,"
    StrSQL = StrSQL & "    dbo.Tblposdata.BoxName , "
    StrSQL = StrSQL & "    dbo.Tblposdata.BoxNamee  "
    StrSQL = StrSQL & " FROM         dbo.TblItemShowDitailses LEFT OUTER JOIN"
    StrSQL = StrSQL & "                      dbo.Tblposdata ON dbo.TblItemShowDitailses.BrnchID = dbo.Tblposdata.BoxID"
    StrSQL = StrSQL & " where id2 = " & val(XPTxtID.text) & " and Type =3 And TblItemShowDitailses.TransType = " & mIndex

    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst

        With Me.grdPos
            .rows = .FixedRows + RsDetails.RecordCount

            For i = .FixedRows To .rows - 1
     
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("POSId")) = RsDetails("brnchid").value & ""
                .TextMatrix(i, .ColIndex("id")) = RsDetails("brnchid").value & ""
                If SystemOptions.UserInterface = EnglishInterface Then
                    .TextMatrix(i, .ColIndex("name")) = RsDetails("BoxNamee").value & ""
                Else
                    .TextMatrix(i, .ColIndex("name")) = RsDetails("BoxName").value & ""
         
                End If
          
                RsDetails.MoveNext
            Next i

        End With

    End If

    '''
    Set RsDetails = Nothing
    '***********************
   
    Set RsDetails = New ADODB.Recordset
    StrSQL = " SELECT     dbo.TblItemGorupShowDitailses.GroupID, dbo.TblItemGorupShowDitailses.Ind, dbo.Groups.GroupName"
    StrSQL = StrSQL & " FROM         dbo.TblItemGorupShowDitailses LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.Groups ON dbo.TblItemGorupShowDitailses.GroupID = dbo.Groups.GroupID"
    StrSQL = StrSQL & " Where (dbo.TblItemGorupShowDitailses.ind = " & val(XPTxtID.text) & ") And TblItemGorupShowDitailses.TransType=" & mIndex
    'StrSQL = StrSQL & " WHERE     (dbo.TblLink_Item_To_Store_Details3.Ind = " & val(XPTxtID.text) & ")"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    For i = 0 To RsDetails.RecordCount - 1
        ListGroupSelected.AddItem IIf(IsNull(RsDetails("GroupName").value), "", RsDetails("GroupName").value)
        ListGroupSelected.ItemData(i) = IIf(IsNull(RsDetails("GroupID").value), "", RsDetails("GroupID").value)
  
        RsDetails.MoveNext
  
    Next i
    
    RsDetails.Close
    Set RsDetails = Nothing
    ' fillapprovData
    ReLineGrid
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub


'Public Sub Retrive2(Optional Lngid As Long = 0)
'    Dim RsDetails As ADODB.Recordset
'    Dim RsDetails1 As ADODB.Recordset
'    Dim RsDetails2 As ADODB.Recordset
   
'
'    Dim i As Integer
'    Dim StrSQL As String
'
'FgBranch.Clear flexClearScrollable, flexClearEverything
'            FgBranch.Rows = 1
'            Me.FgBranch.Enabled = True
'
'            Me.FgItemPloice.Clear flexClearScrollable, flexClearEverything
'            FgItemPloice.Rows = 1
'            Me.FgItemPloice.Enabled = True
'
'              Me.FgItems.Clear flexClearScrollable, flexClearEverything
'            FgItems.Rows = 1
'            Me.FgItems.Enabled = True
'    'On Error GoTo ErrTrap
'    If rs.RecordCount < 1 Then
'        XPTxtCurrent.Caption = 0
'        XPTxtCount.Caption = 0
'        Exit Sub
'    End If
'
'    If rs.EOF Or rs.BOF Then
'        Exit Sub
'    Else
'
'        If Lngid <> 0 Then
'            rs.find "id=" & Lngid, , adSearchForward, adBookmarkFirst
'
'            If rs.EOF Or rs.BOF Then
'                Exit Sub
'            End If
'        End If
'    End If
'
'
'    '''////////
'
'    XPTxtID.text = rs("ID").value
'    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
'     Me.DcbBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
'     Me.TxtNameShow.text = IIf(IsNull(rs("NameShow").value), "", rs("NameShow").value)
'    Startdate.value = IIf(IsNull(rs("StartSDate").value), Date, rs("StartSDate").value)
'    enddate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
'      Me.DcItem1.BoundText = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
'    'Me.DcbGroup.BoundText = IIf(IsNull(rs("GroupID").value), "", rs("GroupID").value)
'     Me.DcbUnit.BoundText = IIf(IsNull(rs("UnitItID").value), "", rs("UnitItID").value)
'     Me.DcbItemDit.BoundText = IIf(IsNull(rs("ItemIDD").value), "", rs("ItemIDD").value)
'     Me.DcbUnitDit.BoundText = IIf(IsNull(rs("UnitItIDD").value), "", rs("UnitItIDD").value)
'     Me.TxtAmountDit.text = IIf(IsNull(rs("AmountD").value), "", rs("AmountD").value)
'     Me.TxtPriceDit.text = IIf(IsNull(rs("PriceD").value), 0, rs("PriceD").value)
'     Me.DcbTypePoliceyDit.ListIndex = IIf(IsNull(rs("TypePoliceD").value), -1, rs("TypePoliceD").value)
'        Me.DcbItemDDis.BoundText = IIf(IsNull(rs("ItemIDDDis").value), "", rs("ItemIDDDis").value)
'           Me.DcbUnitDDis.BoundText = IIf(IsNull(rs("UnitItIDDDis").value), "", rs("UnitItIDDDis").value)
'            Me.TxtAmountDDis.text = IIf(IsNull(rs("AmountDDis").value), "", rs("AmountDDis").value)
'     Me.TxtPriceDDis.text = IIf(IsNull(rs("PriceDDis").value), 0, rs("PriceDDis").value)
'   If rs("AllBranch").value = 0 Then
'   Me.ChAllBranch.value = xtpUnchecked
 '  Else
'   Me.ChAllBranch.value = xtpChecked
'   End If
'DcbBranch1.BoundText = IIf(IsNull(rs("BranchID2").value), "", rs("BranchID2").value)
' If rs("AllPolice").value = 0 Then
'   Me.RdAllPolice.value = False
'   Else
'   Me.RdAllPolice.value = True
'   End If
' If rs("PrivatePolice").value = 0 Then
'   Me.RdPrivatePolice.value = False
'   Else
'   Me.RdPrivatePolice.value = True
'   End If
'
 'Me.DcbtypPolicep.ListIndex = IIf(IsNull(rs("TypePoliceP").value), -1, rs("TypePoliceP").value)
 '
 '  Me.dcbUnitBisc1.BoundText = IIf(IsNull(rs("UnitBisc").value), "", rs("UnitBisc").value)
 '  Me.DcbItemBisc1.BoundText = IIf(IsNull(rs("ItemIDBisc").value), "", rs("ItemIDBisc").value)
 '     Me.TxtAmountBisc1.text = IIf(IsNull(rs("AmountBisc").value), "", rs("AmountBisc").value)
 '      Me.TxtAmountBisc2.text = IIf(IsNull(rs("AmountBisc").value), "", rs("AmountBisc").value)
 '     Me.TxtPriceBisc1.text = IIf(IsNull(rs("PriceBisc").value), "", rs("PriceBisc").value)
 '      Me.TxtPriceBisc2.text = IIf(IsNull(rs("PriceBisc").value), "", rs("PriceBisc").value)
 '     Me.TxtAmountDis.text = IIf(IsNull(rs("AmountDis").value), "", rs("AmountDis").value)
 '
 '   Me.DcbItemDis.BoundText = IIf(IsNull(rs("ItemDis").value), "", rs("ItemDis").value)
 '   Me.dcbUnitDis.BoundText = IIf(IsNull(rs("UnitItDis").value), "", rs("UnitItDis").value)
 '
 'Me.TxtPriceDis.text = IIf(IsNull(rs("PriceDis").value), "", rs("PriceDis").value)
 '     Me.TxtRate.text = IIf(IsNull(rs("Rate").value), "", rs("Rate").value)
 '     Me.TxtRateD.text = IIf(IsNull(rs("RateD").value), "", rs("RateD").value)
 '      Me.DcbUnitGroup.BoundText = IIf(IsNull(rs("UnitGroup").value), "", rs("UnitGroup").value)
''
'''////////
'
''Me.DcbOrderStatus.ListIndex = rs("LinkType").value
'    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
'  '   If IsNull(rs("posted").value) Then
'   '                                                If SystemOptions.UserInterface = ArabicInterface Then
'   '                                                Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
'   '                                              Else
'   '                                                Accredit.Caption = " send to Approval   "
'   '                                            End If
   '                                            Accredit.Enabled = True
'  'Else
'   '                                               If SystemOptions.UserInterface = ArabicInterface Then
'  '                                                  Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
'  ''                                                Else
'   '                                                Accredit.Caption = " sent to Approval   "
'   '                                            End If
'   '                                            Accredit.Enabled = False
'   'End If
'
'
'    Set RsDetails = New ADODB.Recordset
'StrSQL = " select * from TblItemShowDitails where id2 = " & val(XPTxtID.text) & " and Type =0 "
'
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'
'
'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        With Me.FgItems
'       .Rows = .FixedRows + RsDetails.RecordCount
'
'
'        For i = .FixedRows To .Rows - 1
'
'          .TextMatrix(i, .ColIndex("Ser")) = i
'           'TextMatrix(i, .ColIndex("id")) = RsDetails("ItemID").value
'           .TextMatrix(i, .ColIndex("name")) = RsDetails("name").value
'            .TextMatrix(i, .ColIndex("unite")) = RsDetails("unite").value
'          .TextMatrix(i, .ColIndex("group")) = RsDetails("GroupName").value
'
'           .TextMatrix(i, .ColIndex("unite1")) = RsDetails("untgroup").value
'         '   .TextMatrix(i, .ColIndex("amount")) = RsDetails("untgroup").value
'         ' .TextMatrix(i, .ColIndex("price")) = RsDetails("price").value
'         ' .TextMatrix(i, .ColIndex("typedis")) = RsDetails("typedis").value
'         '   .TextMatrix(i, .ColIndex("discount")) = RsDetails("discount").value
'         ' .TextMatrix(i, .ColIndex("itemdis")) = RsDetails("itemdis").value
'         '
'         ' .TextMatrix(i, .ColIndex("unitdis")) = RsDetails("unitdis").value
'         '   .TextMatrix(i, .ColIndex("amountdis")) = RsDetails("amountdis").value
'         ' .TextMatrix(i, .ColIndex("pricedis")) = RsDetails("pricedis").value
'            RsDetails.MoveNext
'        Next i
'End With
'    End If
''''\\\\\\\\\\\\\\\\\\\
'    Set RsDetails = New ADODB.Recordset
'StrSQL = " select * from TblItemShowDitails where id2 = " & val(XPTxtID.text) & " and Type =1 "
'
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'
'
'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        With Me.FgItemPloice
'       .Rows = .FixedRows + RsDetails.RecordCount
'
'
 '       For i = .FixedRows To .Rows - 1
     
 '          .TextMatrix(i, .ColIndex("Ser")) = i
           'TextMatrix(i, .ColIndex("id")) = RsDetails("ItemID").value
 '          .TextMatrix(i, .ColIndex("name")) = RsDetails("name").value
'            .TextMatrix(i, .ColIndex("unite")) = RsDetails("unite").value
 '         .TextMatrix(i, .ColIndex("group")) = RsDetails("GroupName").value
'
'           .TextMatrix(i, .ColIndex("untgroup")) = RsDetails("untgroup").value
'            .TextMatrix(i, .ColIndex("amount")) = RsDetails("amount").value
'          .TextMatrix(i, .ColIndex("price")) = RsDetails("price").value
'          .TextMatrix(i, .ColIndex("typedis")) = RsDetails("typedis").value
 '           .TextMatrix(i, .ColIndex("discount")) = RsDetails("discount").value
'          .TextMatrix(i, .ColIndex("itemdis")) = RsDetails("itemdis").value
'
'          .TextMatrix(i, .ColIndex("unitdis")) = RsDetails("unitdis").value
'            .TextMatrix(i, .ColIndex("amountdis")) = RsDetails("amountdis").value
'          .TextMatrix(i, .ColIndex("pricedis")) = RsDetails("pricedis").value
'            RsDetails.MoveNext
'        Next i
'End With
'    End If
'   '''
'       Set RsDetails = New ADODB.Recordset
'StrSQL = " select * from TblItemShowBranch where id2 = " & val(XPTxtID.text) & " "
'
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'
'
'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        With Me.FgBranch
'       .Rows = .FixedRows + RsDetails.RecordCount
'
'
'        For i = .FixedRows To .Rows - 1
'
'           .TextMatrix(i, .ColIndex("Ser")) = i
'          .TextMatrix(i, .ColIndex("id")) = RsDetails("BranchID").value
'           .TextMatrix(i, .ColIndex("name")) = RsDetails("Namebranch").value
'
'            RsDetails.MoveNext
'        Next i
'End With
'    End If
'   '''
'
'
'     RsDetails.Close
'    Set RsDetails = Nothing
'   ' fillapprovData
'    ReLineGrid
'    XPTxtCurrent.Caption = rs.AbsolutePosition
'    XPTxtCount.Caption = rs.RecordCount
'    Exit Sub
'ErrTrap:
'End Sub
'
'
'Public Sub Retrive3(Optional Lngid As Long = 0)
'    Dim RsDetails As ADODB.Recordset
'    Dim RsDetails1 As ADODB.Recordset
'    Dim RsDetails2 As ADODB.Recordset
'
'
'    Dim i As Integer
'    Dim StrSQL As String
'
'FgBranch.Clear flexClearScrollable, flexClearEverything
'            FgBranch.Rows = 1
'            Me.FgBranch.Enabled = True
'
'            Me.FgItemPloice.Clear flexClearScrollable, flexClearEverything
'            FgItemPloice.Rows = 1
'            Me.FgItemPloice.Enabled = True
'
'              Me.FgItems.Clear flexClearScrollable, flexClearEverything
'            FgItems.Rows = 1
'            Me.FgItems.Enabled = True
'    'On Error GoTo ErrTrap
'    If rs.RecordCount < 1 Then
'        XPTxtCurrent.Caption = 0
'        XPTxtCount.Caption = 0
'        Exit Sub
'    End If
'
'    If rs.EOF Or rs.BOF Then
'        Exit Sub
'    Else
'
'        If Lngid <> 0 Then
'            rs.find "id=" & Lngid, , adSearchForward, adBookmarkFirst
'
'            If rs.EOF Or rs.BOF Then
'                Exit Sub
'            End If
'        End If
'    End If
'
'
'    '''////////
'
'    XPTxtID.text = rs("ID").value
'    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
'     Me.DcbBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
'     Me.TxtNameShow.text = IIf(IsNull(rs("NameShow").value), "", rs("NameShow").value)
'    Startdate.value = IIf(IsNull(rs("StartSDate").value), Date, rs("StartSDate").value)
'    enddate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
'      Me.DcbItem.BoundText = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
'    Me.DcbGroup.BoundText = IIf(IsNull(rs("GroupID").value), "", rs("GroupID").value)
'     Me.DcbUnit.BoundText = IIf(IsNull(rs("UnitItID").value), "", rs("UnitItID").value)
'     Me.DcbItemDit.BoundText = IIf(IsNull(rs("ItemIDD").value), "", rs("ItemIDD").value)
'     Me.DcbUnitDit.BoundText = IIf(IsNull(rs("UnitItIDD").value), "", rs("UnitItIDD").value)
'     Me.TxtAmountDit.text = IIf(IsNull(rs("AmountD").value), "", rs("AmountD").value)
'     Me.TxtPriceDit.text = IIf(IsNull(rs("PriceD").value), 0, rs("PriceD").value)
'     Me.DcbTypePoliceyDit.ListIndex = IIf(IsNull(rs("TypePoliceD").value), -1, rs("TypePoliceD").value)
'        Me.DcbItemDDis.BoundText = IIf(IsNull(rs("ItemIDDDis").value), "", rs("ItemIDDDis").value)
'           Me.DcbUnitDDis.BoundText = IIf(IsNull(rs("UnitItIDDDis").value), "", rs("UnitItIDDDis").value)
'            Me.TxtAmountDDis.text = IIf(IsNull(rs("AmountDDis").value), "", rs("AmountDDis").value)
'     Me.TxtPriceDDis.text = IIf(IsNull(rs("PriceDDis").value), 0, rs("PriceDDis").value)
'   If rs("AllBranch").value = 0 Then
'   Me.ChAllBranch.value = xtpUnchecked
'   Else
'   Me.ChAllBranch.value = xtpChecked
'   End If
'DcbBranch1.BoundText = IIf(IsNull(rs("BranchID2").value), "", rs("BranchID2").value)
' If rs("AllPolice").value = 0 Then
'   Me.RdAllPolice.value = False
'   Else
'   Me.RdAllPolice.value = True
'   End If
' If rs("PrivatePolice").value = 0 Then
'   Me.RdPrivatePolice.value = False
'   Else
'   Me.RdPrivatePolice.value = True
'   End If
'
' Me.DcbtypPolicep.ListIndex = IIf(IsNull(rs("TypePoliceP").value), -1, rs("TypePoliceP").value)
'
'   Me.dcbUnitBisc1.BoundText = IIf(IsNull(rs("UnitBisc").value), "", rs("UnitBisc").value)
'   Me.DcbItemBisc1.BoundText = IIf(IsNull(rs("ItemIDBisc").value), "", rs("ItemIDBisc").value)
'      Me.TxtAmountBisc1.text = IIf(IsNull(rs("AmountBisc").value), "", rs("AmountBisc").value)
'       Me.TxtAmountBisc2.text = IIf(IsNull(rs("AmountBisc").value), "", rs("AmountBisc").value)
'      Me.TxtPriceBisc1.text = IIf(IsNull(rs("PriceBisc").value), "", rs("PriceBisc").value)
'       Me.TxtPriceBisc2.text = IIf(IsNull(rs("PriceBisc").value), "", rs("PriceBisc").value)
'      Me.TxtAmountDis.text = IIf(IsNull(rs("AmountDis").value), "", rs("AmountDis").value)
'
'    Me.DcbItemDis.BoundText = IIf(IsNull(rs("ItemDis").value), "", rs("ItemDis").value)
'    Me.dcbUnitDis.BoundText = IIf(IsNull(rs("UnitItDis").value), "", rs("UnitItDis").value)
'
' Me.TxtPriceDis.text = IIf(IsNull(rs("PriceDis").value), "", rs("PriceDis").value)
'      Me.TxtRate.text = IIf(IsNull(rs("Rate").value), "", rs("Rate").value)
'      Me.TxtRateD.text = IIf(IsNull(rs("RateD").value), "", rs("RateD").value)
'       Me.DcbUnitGroup.BoundText = IIf(IsNull(rs("UnitGroup").value), "", rs("UnitGroup").value)

''////////

'Me.DcbOrderStatus.ListIndex = rs("LinkType").value
'    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
  '   If IsNull(rs("posted").value) Then
   '                                                If SystemOptions.UserInterface = ArabicInterface Then
   '                                                Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
   '                                              Else
   '                                                Accredit.Caption = " send to Approval   "
   '                                            End If
   '                                            Accredit.Enabled = True
  'Else
   '                                               If SystemOptions.UserInterface = ArabicInterface Then
  '                                                  Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
  ''                                                Else
   '                                                Accredit.Caption = " sent to Approval   "
   '                                            End If
   '                                            Accredit.Enabled = False
   'End If
   
   
'    Set RsDetails = New ADODB.Recordset
'StrSQL = " select * from TblItemShowDitails where id2 = " & val(XPTxtID.text) & " and Type =0 "

'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        With Me.FgItems
'       .Rows = .FixedRows + RsDetails.RecordCount


'        For i = .FixedRows To .Rows - 1
     
'          .TextMatrix(i, .ColIndex("Ser")) = i
           'TextMatrix(i, .ColIndex("id")) = RsDetails("ItemID").value
'           .TextMatrix(i, .ColIndex("name")) = RsDetails("name").value
'            .TextMatrix(i, .ColIndex("unite")) = RsDetails("unite").value
'          .TextMatrix(i, .ColIndex("group")) = RsDetails("GroupName").value
'
'           .TextMatrix(i, .ColIndex("unite1")) = RsDetails("untgroup").value
         '   .TextMatrix(i, .ColIndex("amount")) = RsDetails("untgroup").value
         ' .TextMatrix(i, .ColIndex("price")) = RsDetails("price").value
         ' .TextMatrix(i, .ColIndex("typedis")) = RsDetails("typedis").value
         '   .TextMatrix(i, .ColIndex("discount")) = RsDetails("discount").value
         ' .TextMatrix(i, .ColIndex("itemdis")) = RsDetails("itemdis").value
         '
         ' .TextMatrix(i, .ColIndex("unitdis")) = RsDetails("unitdis").value
'         '   .TextMatrix(i, .ColIndex("amountdis")) = RsDetails("amountdis").value
'         ' .TextMatrix(i, .ColIndex("pricedis")) = RsDetails("pricedis").value
'            RsDetails.MoveNext
'        Next i
'End With
'    End If
'''\\\\\\\\\\\\\\\\\\\
'    Set RsDetails = New ADODB.Recordset
'StrSQL = " select * from TblItemShowDitails where id2 = " & val(XPTxtID.text) & " and Type =1 "
'
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


 '   If Not (RsDetails.BOF Or RsDetails.EOF) Then
 '       RsDetails.MoveFirst
 '       With Me.FgItemPloice
 '      .Rows = .FixedRows + RsDetails.RecordCount


 '       For i = .FixedRows To .Rows - 1
 '
 '          .TextMatrix(i, .ColIndex("Ser")) = i
 '          'TextMatrix(i, .ColIndex("id")) = RsDetails("ItemID").value
 '          .TextMatrix(i, .ColIndex("name")) = RsDetails("name").value
 '           .TextMatrix(i, .ColIndex("unite")) = RsDetails("unite").value
 '         .TextMatrix(i, .ColIndex("group")) = RsDetails("GroupName").value
 '
 '          .TextMatrix(i, .ColIndex("untgroup")) = RsDetails("untgroup").value
 '           .TextMatrix(i, .ColIndex("amount")) = RsDetails("amount").value
 '         .TextMatrix(i, .ColIndex("price")) = RsDetails("price").value
 '         .TextMatrix(i, .ColIndex("typedis")) = RsDetails("typedis").value
 '           .TextMatrix(i, .ColIndex("discount")) = RsDetails("discount").value
''          .TextMatrix(i, .ColIndex("itemdis")) = RsDetails("itemdis").value
'
'          .TextMatrix(i, .ColIndex("unitdis")) = RsDetails("unitdis").value
'            .TextMatrix(i, .ColIndex("amountdis")) = RsDetails("amountdis").value
'          .TextMatrix(i, .ColIndex("pricedis")) = RsDetails("pricedis").value
'            RsDetails.MoveNext
'        Next i
'End With
'    End If
   '''
'       Set RsDetails = New ADODB.Recordset
'StrSQL = " select * from TblItemShowBranch where id2 = " & val(XPTxtID.text) & " "
'
'    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'

'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        With Me.FgBranch
'       .Rows = .FixedRows + RsDetails.RecordCount
'

'        For i = .FixedRows To .Rows - 1
'
'           .TextMatrix(i, .ColIndex("Ser")) = i
'          .TextMatrix(i, .ColIndex("id")) = RsDetails("BranchID").value
'           .TextMatrix(i, .ColIndex("name")) = RsDetails("Namebranch").value
'
'            RsDetails.MoveNext
'        Next i
'End With
'    End If
   '''
    
    
   '  RsDetails.Close
   ' Set RsDetails = Nothing
'   ' fillapprovData
'    ReLineGrid
'    XPTxtCurrent.Caption = rs.AbsolutePosition
'    XPTxtCount.Caption = rs.RecordCount
'    Exit Sub
'ErrTrap:
'End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "id='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

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

'Function createVoucher()
'    Dim bankDes As String
'    Dim AccountCode As String
'
'    Dim Employee_account As String
'    Dim NoteID As String
'    Dim sql As String
'
'    '//////////////////////////////////////Notes////////////////////////////////////
'    Dim line_no As Integer
'   Dim RsNotes As New ADODB.Recordset
''    RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'   If Me.TxtModFlg.text = "E" Then
                  
'       sql = "Delete notes where NoteID=" & val(Me.TxtNoteID.text)
'Else
'
'    End If
'
''    RsNotes.AddNew
'   NoteID = CStr(TxtNoteID.text)
'   RsNotes("NoteID").value = CStr(TxtNoteID.text)
'
''   bankDes = "”‰œ «” Õﬁ«ﬁ ⁄„Ê·«     " '& DcComponentType.text & Chr(13)
'
'  bankDes = bankDes & "  „‰ «·› —… " & DtpaFrom.value & "  «·Ï «·› —… : " & DateTo.value
'  RsNotes("NoteType").value = 5151
'  RsNotes("NoteDate").value = XPDtbTrans.value
'  RsNotes("UserID").value = user_id
'  RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '????? ?????
''   RsNotes("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '????? ??? ?????
'  RsNotes("numbering_type").value = sand_numbering_type(0) '??? ????? ??? ?????
'  RsNotes("numbering_type1").value = sand_numbering_type(51) '??? ????? ??? ????????
'  RsNotes("sanad_year").value = year(XPDtbTrans.value)
'  RsNotes("sanad_month").value = Month(XPDtbTrans.value)
'  RsNotes("note_value_by_characters").value = WriteNo(Format(val(lbl(11).Caption), "0.00"), 0, True, ".")
'  'RsNotes("remark").value = TxtRemarks.text & bankDes
'  RsNotes("Branch_no").value = val(Me.Dcbranch.BoundText)
'
'  RsNotes.update
                
'  line_no = 1
'
'    If Fg.Rows > 1 And val(lbl(11).Caption) > 0 Then
'        Dim RsDev  As ADODB.Recordset
'        Set RsDev = New ADODB.Recordset
'        RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'        AccountCode = Account_Code_dynamic
'
'        RsDev.AddNew
'        RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
'        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
'        RsDev("DEV_ID_Line_No").value = line_no '1 'line_no
'        RsDev("DEV_ID_Line_No1").value = setfoxy_Line
'        RsDev("Account_Code").value = AccountCode
'        RsDev("Value").value = Round(val(Me.lbl(11).Caption), 2)
'        RsDev("Credit_Or_Debit").value = 0
                    
'        RsDev("RecordDate").value = Me.XPDtbTrans.value
'        RsDev("Notes_ID").value = val(Me.TxtNoteID.text)   '(XPTxtID.text)
'        RsDev("Double_Entry_Vouchers_Description").value = bankDes   'Fg_Journal.TextMatrix(i, Fg_Journal.ColIndex("des")) '
'        RsDev("Double_Entry_Vouchers_Descriptione").value = bankDes
''
'       RsDev("UserID").value = user_id
'       RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                         
'       RsDev.update
'   End If

' ??????
          
'    If Fg.Rows > 1 And val(lbl(11).Caption) > 0 Then
'
'       Dim i  As Integer
'       Dim LngDevID  As Long
'
'        With Fg
 
'            For i = .FixedRows To .Rows - 1

'                If .TextMatrix(i, .ColIndex("Emp_ID")) <> "" And val(.TextMatrix(i, .ColIndex("net"))) > 0 Then
'                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '???? ????? ??? ????
'                    AccountCode = Employee_account
'
'                    line_no = line_no + 1
'
'                    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'
'                    If ModAccounts.AddNewDev(LngDevID, line_no, AccountCode, Round(val(.TextMatrix(i, .ColIndex("net"))), 2), 1, "" & bankDes & " „‰ «·«„— —ﬁ„ " & .TextMatrix(i, .ColIndex("ID_Aut")), val(NoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , .TextMatrix(i, .ColIndex("net")), , , , bankDes, , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                        GoTo ErrTrap
'
'                    End If
'
'                End If
'
'            Next i

'        End With
    
'    End If

'    updateNotesValueAndNobytext (val(NoteID))

'ErrTrap:

'End Function
Private Sub SaveData()
    Dim st                As String
    Dim astrSplitItems()  As String
    Dim astrSplitItems2() As String
    Dim nElements         As Integer
    Dim j                 As Integer
    Dim RsDetails1        As ADODB.Recordset
 
    Dim Msg               As String
    Dim RsTemp            As New ADODB.Recordset
    Dim StrSQL            As String
    Dim BeginTrans        As Boolean
    Dim RsDetails         As ADODB.Recordset

    Dim i                 As Integer
    Dim LngDevID          As Long
    Dim LngDevLineNo      As Long
    Dim StrAccountCode    As String
    Dim sql               As String
    '   On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If Me.TxtNameShow.text = "" Then
            MsgBox "ÌÃ» ﬂ «»… «”„ «·⁄—÷"
            Me.TxtNameShow.SetFocus
            Exit Sub
        End If

    End If
 
    '    If TxtNoteSerial.text = "" Then
    '        If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
    '            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
    '        Else
    '
    '            If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
    '                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
    '            Else
    '                       TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
    '            End If
    '        End If
    '    End If

    'Dim TxtNoteSerial1str As String

    '    If TxtNoteSerial1.text = "" Then
    '    TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbTrans.value, 51, 5151)
    '
    '                If TxtNoteSerial1str = "error" Then
    '                    MsgBox " ·« Ì„ﬂ‰ «÷«›…     ”‰œ ⁄„Ê·«  „” Õﬁ…  ÃœÌœ… ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
    '                Else
                               
    '                    If TxtNoteSerial1str = "" Then
    '                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—…  «·’Ì«‰…  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
    '                    Else
    '             txtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, DCPreFix.text)
    '                    End If

    '                End If
    '    End If

    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.text = "N" Then

        XPTxtID.text = CStr(new_id("TblItemShows", "ID", "", True, " TransType = " & mIndex) & "")
        
        XPTxtID.text = CStr(new_id("TblItemShows", "ID", "", True))
        '               TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
               
        '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
        '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
           
        'TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)

        '  TxtNoteSerial1 = Voucher_coding(val(my_branch), XPDtbTrans.value, 51, 5151)
            
        rs.AddNew
    ElseIf Me.TxtModFlg.text = "E" Then
        StrSQL = "Delete From TblItemShowDitailses Where ID2=" & val(Me.XPTxtID.text) & " And IsNull(TransType,0) = " & mIndex
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblItemGorupShowDitailses Where Ind=" & val(Me.XPTxtID.text) & " And IsNull(TransType,0) = " & mIndex
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblItemShInfo Where Ind=" & val(Me.XPTxtID.text) & " And IsNull(TransType,0) = " & mIndex
        Cn.Execute StrSQL, , adExecuteNoRecords
             
    End If
      
    rs("ID").value = val(XPTxtID.text)
    rs("TransType").value = val(mIndex)
    rs("RecordDate").value = Me.XPDtbTrans.value
    rs("BranchID").value = IIf(Me.DcbBranch.BoundText = "", Null, Me.DcbBranch.BoundText)
    rs("NameShow").value = IIf(Me.TxtNameShow.text = "", "", Me.TxtNameShow.text)
    rs("StartSDate").value = Me.StartDate.value
    rs("EndDate").value = Me.EndDate.value
    rs("ItemID").value = IIf(Me.DcItem1.BoundText = "", Null, Me.DcItem1.BoundText)

    rs("Sa") = opt_sa.value
    rs("Su") = opt_su.value
    rs("Mo") = opt_mo.value
    rs("Tu") = opt_tu.value
    rs("We") = opt_We.value
    rs("Th") = opt_Th.value
    rs("Fr") = opt_Fr.value

    rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, (DBCboClientName.BoundText))

    'rs("GroupID").value = IIf(Me.DcbGroup.BoundText = "", Null, Me.DcbGroup.BoundText)
    rs("UnitItID").value = IIf(Me.DcbUnit.BoundText = "", Null, Me.DcbUnit.BoundText)
    rs("ItemIDD").value = IIf(Me.DcbItemDit.BoundText = "", Null, Me.DcbItemDit.BoundText)
    rs("UnitItIDD").value = IIf(Me.DcbUnitDit.BoundText = "", Null, Me.DcbUnitDit.BoundText)
    rs("AmountD").value = IIf(Me.TxtAmountDit.text = "", Null, Me.TxtAmountDit.text)
    rs("PriceD").value = IIf(Me.TxtPriceDit.text = "", 0, Me.TxtPriceDit.text)
    rs("TypePoliceD").value = IIf(Me.DcbTypePoliceyDit.ListIndex = -1, Null, Me.DcbTypePoliceyDit.ListIndex)
    rs("ItemIDDDis").value = IIf(Me.DcbItemDDis.BoundText = "", Null, Me.DcbItemDDis.BoundText)
    rs("UnitItIDDDis").value = IIf(Me.DcbUnitDDis.BoundText = "", Null, Me.DcbUnitDDis.BoundText)
    rs("AmountDDis").value = IIf(Me.TxtAmountDDis.text = "", Null, Me.TxtAmountDDis.text)
    rs("PriceDDis").value = IIf(Me.TxtPriceDDis.text = "", 0, Me.TxtPriceDDis.text)
    
    '*******************************
    rs!Period = val(txtPeriod.text)
    rs!PLUS = val(txtPlus.text)
    rs!Min = val(txtMin.text)
    ' rs!purQty = val(txtPurQty.text)
    rs!AvgQtyD = val(txtAvgQtyD.text)
    rs!TotalQtyP = val(txtTotalQtyP.text)
    rs!ResultValue = val(txtResultValue.text)
    '********************************
   
    If Me.ChAllBranch.value = xtpChecked Then
        rs("AllBranch").value = 1
    Else
        rs("AllBranch").value = 0
    End If

    rs("BranchID2").value = val(IIf(Me.DcbBranch1.BoundText = "", 0, Me.DcbBranch1.BoundText))
    rs("UnitG").value = IIf(Me.DcbUnitG.BoundText = "", Null, Me.DcbUnitG.BoundText)

    If Me.RdAllPolice.value = True Then
        rs("AllPolice").value = 1
    Else
        rs("AllPolice").value = 0
    End If

    If Me.RdPrivatePolice.value = True Then
        rs("PrivatePolice").value = 1
    Else
        rs("PrivatePolice").value = 0
    End If

    If Me.XPOptShowType(0).value = True Then
        rs("Selected").value = 1
    ElseIf Me.XPOptShowType(1).value = True Then
        rs("Selected").value = 2
    ElseIf Me.XPOptShowType(2).value = True Then
        rs("Selected").value = 3
    Else
        rs("Selected").value = 0
    End If

    '///////////////////////new 20 11 2016
    rs("Sales").value = IIf(val(Me.TxtSales.text) = 0, Null, Me.TxtSales.text)
    rs("GetFree").value = IIf(val(Me.TxtGetFree.text) = 0, Null, Me.TxtGetFree.text)
    rs("Discount").value = IIf(val(Me.txtDiscount.text) = 0, Null, Me.txtDiscount.text)
  
    rs("FromPrice").value = IIf(Me.CboFromPrice.ListIndex = -1, Null, Me.CboFromPrice.ListIndex)
    '///////////////////////new 20 11 2016
     
    rs("TypePoliceP").value = IIf(Me.DcbtypPolicep.ListIndex = -1, Null, Me.DcbtypPolicep.ListIndex)

    rs("UnitBisc").value = IIf(Me.dcbUnitBisc1.BoundText = "", Null, Me.dcbUnitBisc1.BoundText)
    rs("ItemIDBisc").value = IIf(Me.DcbItemBisc1.BoundText = "", Null, Me.DcbItemBisc1.BoundText)
     
    rs("AmountBisc").value = IIf(Me.TxtAmountBisc1.text = "", Null, Me.TxtAmountBisc1.text)
    rs("PriceBisc").value = IIf(Me.TxtPriceBisc1.text = "", 0, Me.TxtPriceBisc1.text)
    rs("AmountBisc").value = IIf(Me.TxtAmountBisc2.text = "", Null, Me.TxtAmountBisc2.text)
    rs("PriceBisc").value = val(IIf(Me.TxtPriceBisc2.text = "", 0, Me.TxtPriceBisc2.text))
    rs("AmountDis").value = IIf(Me.TxtAmountDis.text = "", Null, Me.TxtAmountDis.text)
              
    rs("ItemDis").value = IIf(Me.DcbItemDis.BoundText = "", Null, Me.DcbItemDis.BoundText)
    rs("UnitItDis").value = IIf(Me.dcbUnitDis.BoundText = "", Null, Me.dcbUnitDis.BoundText)

    rs("PriceDis").value = val(IIf(Me.TxtPriceDis.text = "", 0, Me.TxtPriceDis.text))
    rs("Rate").value = val(IIf(Me.txtRate.text = "", 0, Me.txtRate.text))
    rs("RateD").value = val(IIf(Me.TxtRateD.text = "", 0, Me.TxtRateD.text))
 
    rs("UnitGroup").value = IIf(Me.DcbUnitGroup.BoundText = "", Null, Me.DcbUnitGroup.BoundText)
  
    rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)
    '
    rs.update
    '''''''''/////////////////////////////////
    '''//

    Set RsDetails1 = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblItemShInfo Where (1 = -1)"
    RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
         
    ''//
 
    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblItemShowDitailses Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Me.FgItemPloice

        If .rows > 1 Then

            For i = .FixedRows To .rows - 1

                If .TextMatrix(i, .ColIndex("name")) <> "" Then
      
                    RsDetails.AddNew
                    RsDetails("ID2").value = val(XPTxtID.text)
                    RsDetails("TransType").value = val(mIndex)
      
                    RsDetails("Type").value = 0
                    '
                    RsDetails("InfITemSho").value = .TextMatrix(i, .ColIndex("itemssh"))
                    '
                    RsDetails("uniteId").value = val(.TextMatrix(i, .ColIndex("uniteId")))
                    RsDetails("GropId").value = val(.TextMatrix(i, .ColIndex("GropId")))
                    RsDetails("ItemID").value = val(.TextMatrix(i, .ColIndex("id")))
                    RsDetails("typedisid").value = val(.TextMatrix(i, .ColIndex("typedis")))
                    RsDetails("unitdisid").value = val(.TextMatrix(i, .ColIndex("unitdisid")))
                    RsDetails("amount").value = val(.TextMatrix(i, .ColIndex("amount")))
                    RsDetails("amountdis").value = val(.TextMatrix(i, .ColIndex("amountdis")))
                    RsDetails("discount").value = val(.TextMatrix(i, .ColIndex("discount")))
                    RsDetails("pricedis").value = val(.TextMatrix(i, .ColIndex("pricedis")))
                    RsDetails("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
       
                    RsDetails("ItemDisID").value = val(.TextMatrix(i, .ColIndex("itemdisid")))
                    
                    '***********************************
                    RsDetails!Period = val(.TextMatrix(i, .ColIndex("colPeriod")))
                    RsDetails!PLUS = val(.TextMatrix(i, .ColIndex("colPlus")))
                    RsDetails!Min = val(.TextMatrix(i, .ColIndex("colMin")))
                    RsDetails!purQty = val(.TextMatrix(i, .ColIndex("colPurQty")))
                    RsDetails!AvgQtyD = val(.TextMatrix(i, .ColIndex("colAvgQtyD")))
                    RsDetails!TotalQtyP = val(.TextMatrix(i, .ColIndex("colTotalQtyP")))
                    RsDetails!ResultValue = val(.TextMatrix(i, .ColIndex("colResultValue")))
                    '  RsDetails("pricedis").value = .TextMatrix(i, .ColIndex("pricedis"))
                    RsDetails.update

                    If .TextMatrix(i, .ColIndex("itemssh")) <> "" Then
                        st = .TextMatrix(i, .ColIndex("itemssh"))
                        astrSplitItems = Split(st, "@")
     
                        nElements = UBound(astrSplitItems) - LBound(astrSplitItems)

                        For j = 0 To nElements - 1
                            RsDetails1.AddNew
                            astrSplitItems2 = Split(astrSplitItems(j), "#")
                            RsDetails1("ind").value = val(XPTxtID.text)
                            RsDetails1("TransType").value = val(mIndex)
                            RsDetails1("ID2").value = RsDetails("ID").value
                            RsDetails1("TransType").value = val(mIndex)
                            RsDetails1("ItemID").value = astrSplitItems2(0)
                            RsDetails1("UnitID").value = astrSplitItems2(1)
                            RsDetails1("Qntity").value = astrSplitItems2(2)
                            RsDetails1.update
                        Next j
          
                    End If
         
                End If

            Next i

        End If

    End With

    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblItemShowDitailses Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Me.FgBranch

        If .rows > 1 Then

            For i = .FixedRows To .rows - 1

                If .TextMatrix(i, .ColIndex("name")) <> "" Then
                    RsDetails.AddNew
                                    
                    RsDetails("ID2").value = val(XPTxtID.text)
                    RsDetails("TransType").value = val(mIndex)
                    RsDetails("type").value = 1
                    RsDetails("BrnchID").value = val(.TextMatrix(i, .ColIndex("branchid")))
       
                    RsDetails.update
                    '
                End If

            Next i

        End If

    End With
    '***************
    
    Set RsDetails = New ADODB.Recordset
    StrSQL = "SELECT     *  from dbo.TblItemShowDitailses Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Me.grdPos

        If .rows > 1 Then

            For i = .FixedRows To .rows - 1

                If .TextMatrix(i, .ColIndex("name")) <> "" Then
                    RsDetails.AddNew
                    RsDetails("ID2").value = val(XPTxtID.text)
                    RsDetails("TransType").value = val(mIndex)
                    RsDetails("type").value = 3
                    RsDetails("BrnchID").value = val(.TextMatrix(i, .ColIndex("POSId")))
                    RsDetails.update
                    '
                End If

            Next i

        End If

    End With
    '******************

    ''''
    Set RsDetails = New ADODB.Recordset
    '   RsDetails2.Open "TblLink_Item_To_Store_Details3", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     *  from dbo.TblItemGorupShowDitailses Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
             
    For i = 0 To ListGroupSelected.ListCount - 1
        RsDetails.AddNew
        RsDetails("Ind").value = val(XPTxtID.text)
        ' RsDetails("TransType").value = val(mIndex)
        RsDetails("GroupID").value = val(ListGroupSelected.ItemData(i))
        RsDetails.update
           
    Next i
    
    Cn.CommitTrans
    BeginTrans = False
    RsDetails.Close
     
    '   Set RsDetails = Nothing
    'createVoucher

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    
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
    'End If

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

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
Dim StrSQL1 As String
Dim sql As String
Dim i As Integer
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
     
                rs.delete
                        StrSQL = "Delete TblItemShows where id=" & val(Me.TXTNoteID.text) & " And IsNull(TransType,0) = " & mIndex
        Cn.Execute StrSQL, , adExecuteNoRecords
                
                   StrSQL = "Delete TblItemShowDitailses where id2=" & val(Me.TXTNoteID.text) & " And IsNull(TransType,0) = " & mIndex
        Cn.Execute StrSQL, , adExecuteNoRecords
   
        StrSQL = "Delete From TblItemGorupShowDitailses Where Ind=" & val(TXTNoteID.text) & " And IsNull(TransType,0) = " & mIndex
        Cn.Execute StrSQL, , adExecuteNoRecords
               StrSQL = "Delete From TblItemShInfo Where Ind=" & val(TXTNoteID.text) & " And IsNull(TransType,0) = " & mIndex
        Cn.Execute StrSQL, , adExecuteNoRecords


'                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.text)
'                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
   
                StrSQL1 = "Delete From TblItemShowBranch Where id2=" & val(Me.XPTxtID.text) & " And IsNull(TransType,0) = " & mIndex
                Cn.Execute StrSQL1, , adExecuteNoRecords
              
                    clear_all Me
                        ListGroupSelected.Clear
   ' ListStoreSelected.Clear

                   Me.FgBranch.Clear flexClearScrollable, flexClearEverything
                   FgBranch.rows = 2
                   Me.grdPos.Clear flexClearScrollable, flexClearEverything
                   grdPos.rows = 2
                   Me.FgItemPloice.Clear flexClearScrollable, flexClearEverything
                   FgItemPloice.rows = 2
                    Me.FgItems.Clear flexClearScrollable, flexClearEverything
                   FgItems.rows = 2
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                 
                End If
           ' End If
        End If
   Retrive
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



'Function FillApprovedTable()
' Dim RSApproval  As New ADODB.Recordset
'   Set RSApproval = New ADODB.Recordset
'   Dim currentdate As Date
'   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'

' Dim sql As String
'  Dim rs1 As New ADODB.Recordset
' Dim i As Integer
'    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
 ' sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
''  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
 ' sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
 ' sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
'sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.name & "')"
'sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "
'
'    rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs1.RecordCount > 0 Then
'            currentdate = Now
'            For i = 1 To rs1.RecordCount
'              RSApproval.AddNew
'                RSApproval("ScreenName").value = Me.name
'                RSApproval("levelo").value = IIf(IsNull(rs1("levelo").value), Null, rs1("levelo").value)
'               RSApproval("EmpID").value = IIf(IsNull(rs1("EmpID").value), Null, rs1("EmpID").value)
'                RSApproval("levelorder").value = IIf(IsNull(rs1("levelorder").value), Null, rs1("levelorder").value)
'                 RSApproval("currorder").value = IIf(IsNull(rs1("currorder").value), Null, rs1("currorder").value)
'                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
'                   RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
'                RSApproval("Transaction_Date").value = Date
'
'                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.name), currentdate)
'               RSApproval("SendTime").value = currentdate

''                 If i = 1 Then
  '                      RSApproval("Currcursor").value = 1
 ''                        RSApproval("FromUser").value = user_name
  '              End If
  '
  '              RSApproval.update
  '              rs1.MoveNext
  '          Next i
'
'    End If
    
    

'End Function



'Function fillapprovData()
'Dim Num As Integer
' Dim RsDetails As New ADODB.Recordset
' Dim StrSQL As String
'
'
' StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
'StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
'StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
'StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
'StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
'StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
'StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.name & "')"
'StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"
''
 '   RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
' If Not (RsDetails.EOF Or RsDetails.BOF) Then
''        GRID2.Rows = RsDetails.RecordCount + 1
 

 '       For Num = 1 To RsDetails.RecordCount
 '
 '      GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
 '   If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
 '  GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
 '  Else
 '   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
 ''   End If
  '
  '      GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
  '         If SystemOptions.UserInterface = ArabicInterface Then
  '          GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
  '        Else
  '           GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
  '        End If
  '          If SystemOptions.UserInterface = ArabicInterface Then
  '          GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
  '          Else
  '          GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
  '          End If
  '          GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
  '        GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 '
 
'RsDetails.MoveNext
'If Num = RsDetails.RecordCount Then

'        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
'                                If SystemOptions.UserInterface = ArabicInterface Then
'                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
'                                 Else
'                                       Label11.Caption = "Approved"
'                                 End If
'                            Label11.backcolor = &H80FF80
'        Else
'                             If SystemOptions.UserInterface = ArabicInterface Then
'                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
'                            Else
'                                     Label11.Caption = "Currently required Approve"
'                            End If
'                 Label11.backcolor = &HFFFFC0
'        End If

'End If

  '      Next Num
'Else
' GRID2.Rows = 1
'    End If
'RsDetails.Close

'End Function
Private Sub RemoveGridRowPolice()

    With Me.FgItemPloice

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub RemoveGridRowBr2()

    With Me.grdPos

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub
Private Sub RemoveGridRowBr()

    With Me.FgBranch

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub
'Private Sub RemoveGridRowGr()
'
'    With Me.FgItems
'
'        If .Row <= 0 Then Exit Sub
'        .RemoveItem .Row
'    End With
'
'    ReLineGrid
'End Sub
Function FillMylist()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
  

  sql = " SELECT * from  Groups where GroupID>1"
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListGroupAll.Clear
    ListGroupSelected.Clear

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount

            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupAll.AddItem IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
            Else
                ListGroupAll.AddItem IIf(IsNull(rs("GroupNamee").value), "", rs("GroupNamee").value)
            End If

            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("GroupID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

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

Private Sub ReLineGrid()
    Dim i          As Integer
    Dim IntCounter As Integer
    ' Me.lbl(11).Caption = 0
    IntCounter = 0

    With Me.FgBranch

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
    
            End If

        Next i

    End With
    IntCounter = 0
    With Me.grdPos

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
    
            End If

        Next i
 
    End With

    IntCounter = 0

    With Me.FgItemPloice

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
    
            End If

        Next i
 
    End With

    IntCounter = 0

    With Me.FgItems

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
    
            End If

        Next i
 
    End With
End Sub

'Function FillMylist()
'    Dim sql As String
'    Dim rs As New ADODB.Recordset
'    Dim Balance As Double
'    Dim i As Integer
'    sql = " SELECT * from  TblStore"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'  '  ListStoreall.Clear
'   ' ListStoreSelected.Clear
'
'    If rs.RecordCount > 0 Then
'
'        For i = 1 To rs.RecordCount
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'            '    ListStoreall.AddItem IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
'            Else
'             '   ListStoreall.AddItem IIf(IsNull(rs("StoreNamee").value), "", rs("StoreNamee").value)
'            End If
'
'          '  ListStoreall.ItemData(ListStoreall.NewIndex) = rs("StoreID").value
'            rs.MoveNext
'        Next i
'
'    End If
'
'    rs.Close
'
'    'fil
'
'  sql = " SELECT * from  Groups where GroupID>1"
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    ListGroupAll.Clear
'    ListGroupSelected.Clear
'
'    If rs.RecordCount > 0 Then
'
'        For i = 1 To rs.RecordCount
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'                ListGroupAll.AddItem IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
'            Else
'                ListGroupAll.AddItem IIf(IsNull(rs("GroupNamee").value), "", rs("GroupNamee").value)
'            End If

'            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("GroupID").value
'            rs.MoveNext
'        Next i
'
'    End If
'
'    rs.Close

'End Function
Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "‘«‘…  ⁄—Ê÷ «·«’‰«›   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘…  ⁄—Ê÷ «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " ‘«‘…  ⁄—Ê÷ «·«’‰«›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… ⁄—Ê÷ «·«’‰«›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘…  ⁄—Ê÷ «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… ⁄—Ê÷ «·«’‰«›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… ⁄—Ê÷ «·«’‰«›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘… ⁄—Ê÷ «·«’‰«›  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘…  ⁄—Ê÷ «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘…  ⁄—Ê÷ «·«’‰«› ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "‘«‘…  ⁄—Ê÷ «·«’‰«› ", 1, 15204351, -2147483630
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
       
                'SaveData

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub



 'Private Sub RemoveGridRow()


Private Sub XPOptShowType_Click(Index As Integer)

If XPOptShowType(1).value = True Then
RdPrivatePolice.value = False
Fra(7).Enabled = False
Frame11.Enabled = True
DcbUnitGroup.Enabled = True
Else
Frame11.Enabled = False
DcbUnitGroup.Enabled = False
End If
If XPOptShowType(0).value = True Then
RdPrivatePolice.value = False
DcbUnit.Enabled = True
Else
DcbUnit.Enabled = False
End If
If XPOptShowType(2).value = True Then
Fra(7).Enabled = True
RdPrivatePolice.value = False
DcItem1.Enabled = True
DcbUnitG.Enabled = True
Text1.Enabled = True
Else
Fra(7).Enabled = False
Text1.Enabled = False

DcItem1.Enabled = False
DcbUnitG.Enabled = False
End If


End Sub

