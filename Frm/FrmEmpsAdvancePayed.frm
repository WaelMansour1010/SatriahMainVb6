VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmEmpsAdvancePayed 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "—œ ”·›… „ÊŸ›"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   Icon            =   "FrmEmpsAdvancePayed.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   7020
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "”·›Ì«  «·„ÊŸ›"
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
      Height          =   2895
      Index           =   0
      Left            =   690
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1800
      Width           =   6255
      Begin VB.TextBox TxtPaymentCounts 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   4260
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   37
         Top             =   2490
         Width           =   825
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2205
         Left            =   90
         TabIndex        =   31
         Top             =   270
         Width           =   6135
         _cx             =   10821
         _cy             =   3889
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmEmpsAdvancePayed.frx":038A
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
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·œ›⁄« "
         Height          =   255
         Index           =   12
         Left            =   5220
         TabIndex        =   38
         Top             =   2550
         Width           =   855
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   11
         Left            =   2460
         TabIndex        =   36
         Top             =   2550
         Width           =   855
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·”œ«œ"
         Height          =   255
         Index           =   10
         Left            =   3300
         TabIndex        =   35
         Top             =   2550
         Width           =   825
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   9
         Left            =   180
         TabIndex        =   34
         Top             =   2520
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… —’Ìœ «·„ÊŸ›:"
         Height          =   255
         Index           =   5
         Left            =   900
         TabIndex        =   33
         Top             =   2520
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin VB.TextBox TxtAdvanceValue 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   2
      Top             =   4800
      Width           =   2955
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4500
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   735
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   7035
      _cx             =   12409
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
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   " —œ ”·›… „ÊŸ›"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   1
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1185
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
         ButtonImage     =   "FrmEmpsAdvancePayed.frx":04A9
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
         ButtonImage     =   "FrmEmpsAdvancePayed.frx":0843
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
         TabIndex        =   6
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
         ButtonImage     =   "FrmEmpsAdvancePayed.frx":0BDD
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
         ButtonImage     =   "FrmEmpsAdvancePayed.frx":0F77
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
         Left            =   2880
         Picture         =   "FrmEmpsAdvancePayed.frx":1311
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   2370
      TabIndex        =   8
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   94765057
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DcboEmpName 
      Height          =   315
      Left            =   2880
      TabIndex        =   9
      Top             =   1080
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   30
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5580
      Width           =   6945
      _cx             =   12250
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
         Left            =   6150
         TabIndex        =   11
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
         Left            =   5295
         TabIndex        =   12
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
         Left            =   4455
         TabIndex        =   13
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
         Left            =   3600
         TabIndex        =   14
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
         Left            =   2745
         TabIndex        =   15
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
         Left            =   0
         TabIndex        =   16
         Top             =   75
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
      Begin ImpulseButton.ISButton Cmdprint 
         Height          =   405
         Left            =   855
         TabIndex        =   17
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   714
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   5
         Left            =   1860
         TabIndex        =   32
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
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   3900
      TabIndex        =   18
      Top             =   5160
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   2880
      TabIndex        =   19
      Top             =   1440
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·⁄„·Ì…"
      Height          =   285
      Index           =   4
      Left            =   5910
      TabIndex        =   29
      Top             =   750
      Width           =   1005
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   285
      Index           =   3
      Left            =   5910
      TabIndex        =   28
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»·€"
      Height          =   315
      Index           =   2
      Left            =   5790
      TabIndex        =   27
      Top             =   4785
      Width           =   1005
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   3750
      TabIndex        =   26
      Top             =   750
      Width           =   705
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   6045
      TabIndex        =   25
      Top             =   5235
      Width           =   900
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2430
      TabIndex        =   24
      Top             =   5190
      Width           =   1065
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   690
      TabIndex        =   23
      Top             =   5190
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   22
      Top             =   5220
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   21
      Top             =   5220
      Width           =   615
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   5910
      TabIndex        =   20
      Top             =   1560
      Width           =   1005
   End
End
Attribute VB_Name = "FrmEmpsAdvancePayed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim EmpReport As ClsEmployeeReport
Dim Employee_account As String

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            XPTxtID.text = CStr(new_id("TblEmpAdvance", "AdvanceID", "", True))
            Me.DCboUserName.BoundText = user_id
            'TxtPaymentCounts.text = 1
            XPDtbTrans.SetFocus

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            '        Load FrmNotesSearch
            '        FrmNotesSearch.SearchType = 3
            '        FrmNotesSearch.Show vbModal
        Case 6
            Unload Me

        Case 8
            CalCulateParts
        
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()

End Sub

Private Sub CmdPrint_Click()
    Dim msg As String
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        Set EmpReport = New ClsEmployeeReport
        EmpReport.EmpPyment Fg.TextMatrix(1, Fg.ColIndex("AdvanceID")) 'XPTxtID.text
    
    Else
        msg = "⁄„·Ì… «·ÿ»«⁄… €Ì— „ «Õ… Õ«·Ì«"
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub DcboEmpName_Change()
    GetEmpAdv
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
    GetEmpAdv
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    CalCulateParts
End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    If Col = Fg.ColIndex("Checked") Then
        Cancel = False
    Else
        Cancel = True
    End If

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        .RowHeightMin = 300
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmdprint.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.DcboEmpName
    SetDtpickerDate Me.XPDtbTrans
    YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEmpAdvance where (TblEmpAdvance.AdvanceType =1) Order By AdvanceID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
    Retrive
    Me.TxtModFlg.text = "R"

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
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

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmdprint.Caption = "Print"
    Cmd(6).Caption = "Exit"

    Lbl(1).Caption = "Date"

    With Me.Fg
        .TextMatrix(0, .ColIndex("Checked")) = "select"
        .TextMatrix(0, .ColIndex("AdvanceID")) = "AdvanceID"
        .TextMatrix(0, .ColIndex("PartNO")) = "PartNO"
        .TextMatrix(0, .ColIndex("PartValue")) = "PartValue"
        .TextMatrix(0, .ColIndex("PartDate")) = "PartDate"

    End With

    Me.Caption = "Employees Advance Payed"
    EleHeader.Caption = "Emp Advance Payed"
    Lbl(4).Caption = "Code"
    Lbl(3).Caption = "Name"
    Lbl(2).Caption = "value"
    Lbl(0).Caption = "Box"
    Lbl(3).Caption = "Name"
    Fra(0).Caption = "Employee advances"
    Lbl(12).Caption = "Count"
    Lbl(10).Caption = "Payed"
    Lbl(5).Caption = "Balance"
    Lbl(8).Caption = "By"
    Lbl(7).Caption = "Curr rec."
    Lbl(6).Caption = "Rec. Count"

End Sub

Private Sub YearMonth()

    'Dim I As Integer
    'Dim IntDefIndex As Integer
    '
    'CmbMonth.Clear
    'For I = 1 To 12
    '    CmbMonth.AddItem MonthName(I)
    'Next
    'CmbMonth.ListIndex = Month(Date) - 1
    'CboYear.Clear
    'For I = 2000 To 2050
    '    CboYear.AddItem I
    '    If I = Year(Date) Then
    '        IntDefIndex = CboYear.NewIndex
    '    End If
    'Next
    'CboYear.ListIndex = IntDefIndex
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

Private Sub ImgFavorites_Click()
AddTofaforites Me.name, Me.Caption, Me.Caption

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
            TxtAdvanceValue.locked = True
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
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
            TxtAdvanceValue.locked = False
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
            TxtAdvanceValue.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
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
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.find "AdvanceID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.text = IIf(IsNull(rs("AdvanceID").value), "", val(rs("AdvanceID").value))
    XPDtbTrans.value = IIf(IsNull(rs("AdvanceDate").value), Date, rs("AdvanceDate").value)
    Me.DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    Me.TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), "", rs("PaymentCounts").value)

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Set RsDetails = New ADODB.Recordset
    'StrSQL = "Select * From  TblEmpAdvanceDetails Where AdvanceID=" & Val(XPTxtID.text)
    StrSQL = "Select * From  TblEmpAdvanceDetails Where OrgTableID=" & val(XPTxtID.text)
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.Rows = Fg.FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        Fg.Rows = Fg.FixedRows + RsDetails.RecordCount

        For i = Me.Fg.FixedRows To Fg.Rows - 1
            Fg.TextMatrix(i, Fg.ColIndex("AdvanceID")) = RsDetails("AdvanceID").value
            Fg.TextMatrix(i, Fg.ColIndex("PartNO")) = RsDetails("PartNO").value
            Fg.TextMatrix(i, Fg.ColIndex("PartValue")) = RsDetails("PartValue").value
            Fg.TextMatrix(i, Fg.ColIndex("OrgTableID")) = RsDetails("OrgTableID").value
            Fg.TextMatrix(i, Fg.ColIndex("TableID")) = RsDetails("TableID").value
          
            Fg.TextMatrix(i, Fg.ColIndex("PartDate")) = DisplayDate(CDate(RsDetails("PartDate").value))
            Fg.Cell(flexcpChecked, i, Fg.ColIndex("Checked")) = flexChecked
            RsDetails.MoveNext
        Next i

    End If

    RsDetails.Close
    Set RsDetails = Nothing
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
            msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
            MsgBox msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboEmpName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If Trim(Me.DcboBox.BoundText) = "" Then
            msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
            MsgBox msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboBox.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
   
        CalCulateParts
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then
            XPTxtID.text = CStr(new_id("TblEmpAdvance", "AdvanceID", "", True))
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From TblEmpAdvanceDetails Where AdvanceID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
    
        rs("AdvanceID").value = val(XPTxtID.text)
        rs("AdvanceDate").value = XPDtbTrans.value
        rs("Emp_ID").value = Me.DcboEmpName.BoundText
        rs("AdvanceValue").value = IIf(TxtAdvanceValue.text = "", Null, val(TxtAdvanceValue.text))
        rs("BoxID").value = IIf(Me.DcboBox.BoundText = "", Null, Me.DcboBox.BoundText)
        '   rs("PaymentCounts").value = Val(Me.TxtPaymentCounts.text)
        '   rs("AutoDiscount").value = Null
        '   rs("FirstMonthPayment").value = Null
        '   rs("FirstYearPayment").value = Null
        rs("UserID").value = Me.DCboUserName.BoundText
        rs("AdvanceType").value = 1
        rs("RetrunID").value = Null
        rs.update
        Set RsDetails = New ADODB.Recordset
        RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        For i = Me.Fg.FixedRows To Fg.Rows - 1

            If Fg.Cell(flexcpChecked, i, Fg.ColIndex("Checked")) = flexChecked Then
                ' RsDetails.AddNew
                '     RsDetails("AdvanceID").value = Val(XPTxtID.text)
                '     RsDetails("PartNO").value = Fg.TextMatrix(i, Fg.ColIndex("PartNO"))
                '     RsDetails("PartValue").value = Fg.TextMatrix(i, Fg.ColIndex("PartValue"))
                '     RsDetails("PartDate").value = Fg.TextMatrix(i, Fg.ColIndex("PartDate"))
                '     RsDetails("OrgTableID").value = Val(XPTxtID.text) ' Fg.TextMatrix(i, Fg.ColIndex("TableID"))
                ' RsDetails.update
                StrSQL = "Update TblEmpAdvanceDetails Set Payed=1 , OrgTableID =" & val(XPTxtID.text) & " Where TableID=" & val(Fg.TextMatrix(i, Fg.ColIndex("TableID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            Else
                StrSQL = "Update TblEmpAdvanceDetails Set Payed=Null , OrgTableID = Null Where TableID=" & val(Fg.TextMatrix(i, Fg.ColIndex("TableID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If

        Next i

        If detect_employee_work_type = 1 Then
            msg = "—œ ”·› „ÊŸ›Ì‰ —ﬁ„ " & val(Me.XPTxtID.text)
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
            StrAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))

            If ModAccounts.AddNewDev(LngDevID, 1, StrAccountCode, val(Me.Lbl(11).Caption), 0, msg, , , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text)) = False Then
                GoTo ErrTrap
            End If
        
            Employee_account = get_EMPLOYEE_Account(val(Me.DcboEmpName.BoundText), "Account_Code")
            StrAccountCode = Employee_account
        
            'StrAccountCode = "a1a3a4" –„„ «·„ÊŸ›Ì‰
            If ModAccounts.AddNewDev(LngDevID, 2, StrAccountCode, val(Me.Lbl(11).Caption), 1, msg, , , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text)) = False Then
                GoTo ErrTrap
            End If
        End If

        Cn.CommitTrans
        BeginTrans = False
        RsDetails.Close
        Set RsDetails = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
                msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                msg = msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End Select

        TxtModFlg.text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        msg = msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        msg = msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "NoteID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

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
    Dim msg As String
    Dim i As Integer
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        msg = msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
    
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords

                ' For i = 1 To Fg.Rows - 1
                '     StrSQL = "Update TblEmpAdvanceDetails Set Payed=0 Where TableID=" & Val(Fg.TextMatrix(i, Fg.ColIndex("OrgTableID"))) & ""
                '     Cn.Execute StrSQL, , adExecuteNoRecords
                ' Next
                For i = Me.Fg.FixedRows To Fg.Rows - 1
    
                    StrSQL = "Update TblEmpAdvanceDetails Set Payed=Null , OrgTableID = Null Where TableID=" & val(Fg.TextMatrix(i, Fg.ColIndex("TableID"))) & ""
                    Cn.Execute StrSQL, , adExecuteNoRecords
         
                Next i
    
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
        msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
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
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "”·› «·„ÊŸ›Ì‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmdprint, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
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

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub TxtAdvanceValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtAdvanceValue.text, 0)
End Sub

Private Sub CalCulateParts()
    Dim i As Integer
    Dim IntCount As Integer

    Dim SngTotal As Single

    With Me.Fg
        SngTotal = 0
        IntCount = 0

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("Checked")) = flexChecked Then
                IntCount = IntCount + 1
                SngTotal = SngTotal + val(.TextMatrix(i, .ColIndex("PartValue")))
            End If

        Next i

    End With

    Me.TxtPaymentCounts.text = IntCount
    Me.Lbl(11).Caption = SngTotal
    TxtAdvanceValue.text = SngTotal
End Sub

Private Sub GetEmpAdv()
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer

    If Me.DcboEmpName.BoundText = "" Then
        Me.Fg.Rows = Fg.FixedRows
        Me.Lbl(9).Caption = 0
    Else
        StrSQL = "SELECT dbo.TblEmpAdvance.Emp_ID, dbo.TblEmpAdvanceDetails.TableID,dbo.TblEm" & "pAdvanceDetails.PartDate, dbo.TblEmpAdvanceDetails.PartValue," & "dbo.TblEmpAdvanceDetails.PartNO, dbo.TblEmpAdvanceDetails.AdvanceID ,dbo.TblEmpAdvance.AdvanceValue FROM   " & "dbo.TblEmpAdvance INNER JOIN dbo.TblEmpAdvanceDetails" & " ON dbo.TblEmpAdvance.AdvanceID = dbo.TblEmpAdvanceDetails.AdvanceID"
        StrSQL = StrSQL + " Where dbo.TblEmpAdvance.Emp_ID=" & val(Me.DcboEmpName.BoundText) & ""
        StrSQL = StrSQL + " AND (dbo.TblEmpAdvanceDetails.Payed IS NULL) "
    
        StrSQL = StrSQL + " Order By dbo.TblEmpAdvanceDetails.TableID "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            '        AdvanceValue
            '      TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
     
            rs.MoveFirst
            Me.Fg.Rows = Fg.FixedRows + rs.RecordCount

            For i = 1 To rs.RecordCount

                With Me.Fg
                    .TextMatrix(i, .ColIndex("TableID")) = IIf(IsNull(rs("TableID").value), "", rs("TableID").value)
                    .TextMatrix(i, .ColIndex("AdvanceID")) = IIf(IsNull(rs("AdvanceID").value), "", rs("AdvanceID").value)
                    .TextMatrix(i, .ColIndex("PartDate")) = IIf(IsNull(rs("PartDate").value), "", rs("PartDate").value)
                    .TextMatrix(i, .ColIndex("PartValue")) = IIf(IsNull(rs("PartValue").value), "", rs("PartValue").value)
                    .TextMatrix(i, .ColIndex("PartNO")) = IIf(IsNull(rs("PartNO").value), "", rs("PartNO").value)
                    '.TextMatrix(I, .ColIndex("")) = IIf(IsNull(Rs("").Value), "", Rs("").Value)
                End With

                rs.MoveNext
            Next i

        Else
            Me.Fg.Rows = Fg.FixedRows
            Me.Lbl(9).Caption = 0
            TxtAdvanceValue.text = ""
        End If
    End If

End Sub
