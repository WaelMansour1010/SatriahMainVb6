VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCase1 
   Caption         =   "«’œ«— «ﬁ”«ÿ «·«Â·«ﬂ"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14190
   Icon            =   "FrmCase1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   14190
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14190
      _cx             =   25030
      _cy             =   11880
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
      Align           =   5
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
      Begin VB.TextBox XPTxtID 
         Height          =   285
         Left            =   7080
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   7440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChkForAllAssets 
         Alignment       =   1  'Right Justify
         Caption         =   "·ﬂ· «·›—Ê⁄"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox TxtRemarks 
         Alignment       =   1  'Right Justify
         Height          =   615
         Left            =   240
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   960
         Width           =   8535
      End
      Begin VB.TextBox TxtFixedAssetInstallmentsid 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Text            =   " "
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Text            =   "R"
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtnoteid 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   11400
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   6360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox TxtValue 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Text            =   "0"
         Top             =   5880
         Width           =   1575
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   675
         Index           =   0
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   13995
         _cx             =   24686
         _cy             =   1191
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   20.25
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
         Caption         =   "«’œ«— «ﬁ”«ÿ «·«Â·«ﬂ"
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
            Index           =   1
            Left            =   1680
            TabIndex        =   10
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
            ButtonImage     =   "FrmCase1.frx":000C
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
            TabIndex        =   11
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
            ButtonImage     =   "FrmCase1.frx":03A6
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
            Index           =   0
            Left            =   1155
            TabIndex        =   12
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
            ButtonImage     =   "FrmCase1.frx":0740
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
            ButtonImage     =   "FrmCase1.frx":0ADA
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   8910
         TabIndex        =   14
         Top             =   6315
         Width           =   795
         _ExtentX        =   1402
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
         Left            =   8400
         TabIndex        =   15
         Top             =   5835
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
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
         Left            =   8115
         TabIndex        =   16
         Top             =   6315
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "≈’œ«—"
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
         Left            =   7245
         TabIndex        =   17
         Top             =   6315
         Width           =   795
         _ExtentX        =   1402
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
         Left            =   5880
         TabIndex        =   18
         Top             =   6315
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "≈«·€«¡ «·«’œ«—"
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
         Left            =   4080
         TabIndex        =   19
         Top             =   6315
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   6360
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcBranch 
         Height          =   315
         Left            =   10320
         TabIndex        =   21
         Top             =   1560
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DpRecordDate 
         Height          =   345
         Left            =   10320
         TabIndex        =   22
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   241893377
         CurrentDate     =   41640
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   8
         Left            =   7200
         TabIndex        =   23
         Top             =   6720
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "«’œ«—"
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
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   3060
         Left            =   240
         TabIndex        =   24
         Top             =   2400
         Width           =   13800
         _cx             =   24342
         _cy             =   5397
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
         Rows            =   10
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCase1.frx":0E74
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
         Height          =   375
         Index           =   5
         Left            =   5640
         TabIndex        =   25
         Top             =   6720
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "«·€«¡ «·«’œ«—"
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   795
         Index           =   3
         Left            =   3360
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5385
         _cx             =   9499
         _cy             =   1402
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
         Caption         =   "≈Œ Ì«— «· «—ÌŒ"
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   1
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
         Begin VB.ComboBox CboYear 
            Height          =   315
            Left            =   765
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   240
            Width           =   1545
         End
         Begin VB.ComboBox CmbMonth 
            Height          =   315
            Left            =   3285
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   240
            Width           =   1305
         End
         Begin VB.CommandButton CmdView 
            Caption         =   "⁄—÷"
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "”‰…"
            Height          =   225
            Index           =   2
            Left            =   2460
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   600
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‘Â—"
            Height          =   210
            Index           =   1
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   525
         End
      End
      Begin MSDataListLib.DataCombo DCGroups 
         Height          =   315
         Left            =   10320
         TabIndex        =   32
         Top             =   1920
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   11
         Left            =   4920
         TabIndex        =   33
         Top             =   6315
         Width           =   915
         _ExtentX        =   1614
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„·«ÕŸ« "
         Height          =   315
         Index           =   0
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «·”Ã· «·Õ«·Ì:"
         Height          =   315
         Index           =   7
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   5880
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œœ «·”Ã·« :"
         Height          =   315
         Index           =   6
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   5880
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label XPTxtCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   5910
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label XPTxtCurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   2370
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   5910
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label LngDevID 
         Height          =   375
         Left            =   7080
         TabIndex        =   43
         Top             =   7680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "—ﬁ„ «·ﬁÌœ"
         Height          =   375
         Left            =   8280
         TabIndex        =   42
         Top             =   7800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ… : "
         Height          =   315
         Index           =   5
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   6360
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " „Ã„Ê⁄Â"
         Height          =   315
         Index           =   14
         Left            =   12720
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1920
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "·›—⁄ „Õœœ"
         Height          =   315
         Index           =   15
         Left            =   12720
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·⁄„·Ì…"
         Height          =   375
         Index           =   17
         Left            =   12720
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·ﬁ”ÿ"
         Height          =   315
         Index           =   3
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„"
         Height          =   375
         Index           =   4
         Left            =   12840
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   " —ﬁ„ «·ﬁÌœ"
         Height          =   255
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "«·«Ã„«·Ì"
         Height          =   255
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   5880
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCase1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim rs As ADODB.Recordset

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long
 
Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Function GetCarIDBy2(Optional Emp_id As Double) As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     id"
sql = sql & " From dbo.TblCarsData"
sql = sql & " Where (Emp_id = " & Emp_id & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetCarIDBy2 = IIf(IsNull(rs2("id").value), 0, rs2("id").value)
Else
GetCarIDBy2 = 0
End If
End Function

Function GetCarIDByEmpID(Optional Emp_id As Double) As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     fixedAssetid"
sql = sql & " From dbo.TblCarsData"
sql = sql & " Where (Emp_id = " & Emp_id & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetCarIDByEmpID = IIf(IsNull(rs2("fixedAssetid").value), 0, rs2("fixedAssetid").value)
Else
GetCarIDByEmpID = 0
End If
End Function


Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "„”·”· " & TxtFixedAssetInstallmentsid.text & CHR(13) & "   «· «—ÌŒ " & DPRecorddate & CHR(13) & "   ‘Â—  " & CmbMonth & CHR(13) & "   ”‰…  " & CboYear & CHR(13) & "   «·›—⁄ " & dcBranch & CHR(13) & "   «·„Ã„Ê⁄Â " & dcGroups & CHR(13) & "   „·«ÕŸ«  " & txtRemarks

    If ChkForAllAssets.value = Checked Then
        LogTextA = LogTextA & CHR(13) & "·ﬂ· «·«’Ê· "
    End If
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Serial " & TxtFixedAssetInstallmentsid.text & CHR(13) & "   Dwat " & DPRecorddate & CHR(13) & "  Month  " & CmbMonth & CHR(13) & "   Year  " & CboYear & CHR(13) & "   Branch " & dcBranch & CHR(13) & "   Group " & dcGroups & CHR(13) & "   Remarks " & txtRemarks

    If ChkForAllAssets.value = Checked Then
        LogTexte = LogTexte & CHR(13) & "For All F.A."
    End If
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 90, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TxtNoteSerial.text, TxtFixedAssetInstallmentsid
    Else
        AddToLogFile CInt(user_id), 90, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtNoteSerial.text, TxtFixedAssetInstallmentsid
    End If
    
End Function

Private Sub ALLButton1_Click()
    FrmShowCol1.show
End Sub

Function check_previous_dev(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from notes where salary=" & year & Month
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev = False
    Else
        check_previous_dev = True
    End If
 
End Function

Function check_previous_dev1(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev1 = False
    Else
        check_previous_dev1 = True
    End If
 
End Function

Function Create_dev()
   
End Function

Function Create_dev1()
End Function

Function ViewInstallmentInformations()
    Dim GroupID As Integer
    Dim BranchID  As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim Msg  As String

    If ChkForAllAssets.Enabled = False Then Exit Function
    If ChkForAllAssets.value = vbUnchecked Then '›Ì Õ«·… ⁄œ„  ÕœÌœ ﬂ· «·«’Ê·
        If Trim(Me.dcBranch.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·›—⁄ ·«‰ﬂ ·„  Õœœ ﬂ·  «·«’Ê·..!!"
            Else
                Msg = "Select Branch Firstly or Check All Assets Check Box..!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            dcBranch.SetFocus
            Sendkeys "{F4}"
            Exit Function
        End If

        GroupID = val(Me.dcGroups.BoundText)
        BranchID = val(Me.dcBranch.BoundText)
        sql = "Select *  From FixedAssets where  Branch_NO = " & BranchID & " And HaveDepreciation = 1 And Status_id = 0 "  ' ·›—⁄ „Õœœ  '«Œ Ì«— ﬂ· „‰ ·Â «Â·«ﬂ ÊÕ«· … Ã«—Ì «·«Â·«ﬂ

 
        If GroupID <> 0 Then
            sql = "Select *  From FixedAssets where group_id=" & GroupID & " and  Branch_NO = " & BranchID & " And HaveDepreciation = 1 And Status_id = 0 "  ' ·›—⁄ „Õœœ Ê„Ã„Ê⁄Â „Õœœ… '«Œ Ì«— ﬂ· „‰ ·Â «Â·«ﬂ ÊÕ«· … Ã«—Ì «·«Â·«ﬂ

        End If

    Else
        sql = "Select *  From FixedAssets where HaveDepreciation=1 and Status_id =0"  '«Œ Ì«— ﬂ· „‰ ·Â «Â·«ﬂ ÊÕ«· … Ã«—Ì «·«Â·«ﬂ
    End If

    sql = sql & " and PurchasePrice>0" ' «· «ﬂœ „‰  ”ÃÌ· ›« Ê—… «·«’·
    sql = sql & " and StartDepreciationDate<=" & SQLDate(DPRecorddate.value, True)    ' «· «ﬂœ „‰      «—ÌŒ »œ«Ì… «·«Â·«ﬂ
  '  sql = sql & " and id = 103"
    '
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1
  'sql = sql & " and  id =149"
'sql = sql & " and id in(186,187,217)"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim AccDepreciation As Double
    Dim KhordaPrice As Double
    Dim RemianInstallments As Double
    Dim CurrentInstalmentNo As Double
    Dim Installmentvalue As Double
    Dim NewAccDepreciation As Double
    Dim FixedAsssetid As Integer
    Dim purchaseprice As Double
    Dim FixedAssetName As String
    Dim fullcode As String
    Dim branch_no As Integer
    Dim DepitAccount As String
   Dim CreditAccount As String
   Dim currentvalue As Double
   Dim group_id As Integer
    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
            FixedAsssetid = val(rs("id"))
            
            branch_no = val(rs("Branch_NO"))
            GetFixedAssetHistory FixedAsssetid, AccDepreciation, RemianInstallments, CurrentInstalmentNo, Installmentvalue, NewAccDepreciation, purchaseprice, FixedAssetName, currentvalue, fullcode, KhordaPrice, group_id, DepitAccount, CreditAccount

            If RemianInstallments > 0 And currentvalue > 0 Then
                AddNewRow FixedAsssetid, RemianInstallments - 1, CurrentInstalmentNo + 1, Installmentvalue, NewAccDepreciation, purchaseprice, FixedAssetName, fullcode, branch_no, group_id, DepitAccount, CreditAccount, KhordaPrice
            ElseIf RemianInstallments = 0 And Round(currentvalue, 0) > 0 Then
                AddNewRow FixedAsssetid, RemianInstallments - 1, CurrentInstalmentNo + 1, currentvalue, NewAccDepreciation, purchaseprice, FixedAssetName, fullcode, branch_no, group_id, DepitAccount, CreditAccount, KhordaPrice
            End If
             If Round(currentvalue, 0) < 0 Then
                currentvalue = currentvalue
             End If
             If Round(NewAccDepreciation, 0) < 0 Then
                NewAccDepreciation = NewAccDepreciation
             End If
             
            rs.MoveNext
        Next i

    End If

    ReLineGrid

End Function

Private Sub AddNewRow(FixedassetId As Integer, _
                      RemianInstallments As Double, _
                      CurrentInstalmentNo As Double, _
                      Installmentvalue As Double, _
                      NewAccDepreciation As Double, _
                      purchaseprice As Double, _
                      FixedAssetName As String, _
                      Optional fullcode As String, Optional branch_no As Integer, _
                      Optional group_id As Integer, Optional DepitAccount As String, Optional CreditAccount As String, Optional KhordaPrice As Double)
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
 
    Me.Grid.rows = Me.Grid.rows + 1
    LngRow = Me.Grid.rows - 1
 
    On Error Resume Next

    With Me.Grid
 .TextMatrix(LngRow, .ColIndex("group_id")) = group_id
  .TextMatrix(LngRow, .ColIndex("DepitAccount")) = DepitAccount
   .TextMatrix(LngRow, .ColIndex("CreditAccount")) = CreditAccount
   
        .TextMatrix(LngRow, .ColIndex("FixedAssetID")) = FixedassetId
        .TextMatrix(LngRow, .ColIndex("Fullcode")) = fullcode
        '
     If RemianInstallments < 0 Then
    '    NewAccDepreciation = Installmentvalue
     End If
        .TextMatrix(LngRow, .ColIndex("FixedAssetName")) = FixedAssetName
        If RemianInstallments < 0 Then
            .TextMatrix(LngRow, .ColIndex("CurrentValue")) = purchaseprice - (Installmentvalue) - KhordaPrice
        Else
            .TextMatrix(LngRow, .ColIndex("CurrentValue")) = purchaseprice - NewAccDepreciation - KhordaPrice
        End If
        .TextMatrix(LngRow, .ColIndex("CurrentValue")) = purchaseprice - NewAccDepreciation - KhordaPrice
        
        .TextMatrix(LngRow, .ColIndex("InstallmentID")) = CurrentInstalmentNo
        .TextMatrix(LngRow, .ColIndex("InstallmentDate")) = DPRecorddate.value
    
       ' If .TextMatrix(LngRow, .ColIndex("CurrentValue")) <= Installmentvalue Then
       '     .TextMatrix(LngRow, .ColIndex("InstallmentValue")) = val(.TextMatrix(LngRow, .ColIndex("CurrentValue")))
       ' Else
       '     .TextMatrix(LngRow, .ColIndex("InstallmentValue")) = Installmentvalue
       ' End If
       If .TextMatrix(LngRow, .ColIndex("CurrentValue")) <= Installmentvalue Then
            .TextMatrix(LngRow, .ColIndex("InstalVal")) = val(.TextMatrix(LngRow, .ColIndex("CurrentValue")))
        Else
            .TextMatrix(LngRow, .ColIndex("InstalVal")) = Installmentvalue
        End If
        
        'salim06082020
        '.TextMatrix(LngRow, .ColIndex("InstalVal")) = Installmentvalue
        'salim06082020
        
        .TextMatrix(LngRow, .ColIndex("AddValue")) = GetAddValue(val(Me.CboYear.text), val(Me.CmbMonth.ListIndex) + 1, FixedassetId)
        .TextMatrix(LngRow, .ColIndex("InstallmentValue")) = val(.TextMatrix(LngRow, .ColIndex("AddValue"))) + val(.TextMatrix(LngRow, .ColIndex("InstalVal")))
        .TextMatrix(LngRow, .ColIndex("AccDepreciation")) = NewAccDepreciation
        .TextMatrix(LngRow, .ColIndex("RemainInstallments")) = RemianInstallments
        .TextMatrix(LngRow, .ColIndex("Branch_NO")) = branch_no
        
    End With
 
End Sub
  Function GetAddValue(Optional YerID As Integer, Optional MothID As Integer, Optional Fixed As Integer) As Double
   Dim sql As String
   Dim Rs4 As ADODB.Recordset
   Set Rs4 = New ADODB.Recordset
 sql = " SELECT     SUM(QstIncValue) AS SmQstIncValue, YEAR(SatrtDate) AS YerID, MONTH(SatrtDate) AS MonthID, FixedID, Distrbute"
 sql = sql & "  From dbo.TblAdditionsAssest"
 sql = sql & "  WHERE     (Distrbute = 0) AND (FixedID = " & Fixed & ") "
 ''AND (MONTH(SatrtDate) >= " & MothID & ") AND (YEAR(SatrtDate) >= " & YerID & ")"
 sql = sql & " GROUP BY YEAR(SatrtDate), MONTH(SatrtDate), FixedID, Distrbute"
 Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs4.RecordCount > 0 Then
 GetAddValue = IIf(IsNull(Rs4("SmQstIncValue").value), 0, Rs4("SmQstIncValue").value)
 Else
 GetAddValue = 0
 End If
  End Function
Private Sub CboYear_Click()
    On Error Resume Next
    Dim str As String
    str = "01/" & CmbMonth.ListIndex + 1 & "/" & CboYear.text
    str = CboYear.text & "-" & CmbMonth.ListIndex + 1 & "-01"

    DPRecorddate.value = MonthLastDay(CDate(str))
    'CmdView_Click
End Sub

Private Sub ChkForAllAssets_Click()
 
    'CmdView_Click
 
End Sub

Private Sub CmbMonth_Click()
    On Error Resume Next
    Dim str As String
    str = "01/" & CmbMonth.ListIndex + 1 & "/" & CboYear.text

    DPRecorddate.value = MonthLastDay(CDate(str))
    'CmdView_Click
    'CmdView_Click
    'If CheckLastInstallmentDate(Me.CmbMonth.ListIndex, Me.CboYear.ListIndex) = True Then
    'ViewInstallmentInformations
    'End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()

End Sub

Function create_report_data()

End Function

Private Sub CmdPrint_Click()
End Sub

Private Sub Combo1_Click()
 
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim cProgress As ClsProgress
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

  On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If Me.CmbMonth.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "«Œ — «·‘Â— «Ê·«"
            Else
                MsgBox "    Specify Month"
            End If

            Exit Sub
        End If

        If Me.CboYear.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "«Œ — «·”‰… «Ê·«"
            Else
                MsgBox "    Specify Year"
            End If

            Exit Sub
        End If
        
    End If
'DPRecorddate.value = Date
    '-------------------------------------------------------------------------------------------
   
    If TxtNoteSerial.text = "" Then
        If Notes_coding(val(Me.dcBranch.BoundText), Me.DPRecorddate.value) = "error" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number ": Exit Sub
            End If

        ElseIf Notes_coding(val(Me.dcBranch.BoundText), DPRecorddate.value) = "" Then

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                MsgBox "You must Define JE Coding ": Exit Sub
            End If

        Else
            TxtNoteSerial.text = Notes_coding(val(Me.dcBranch.BoundText), DPRecorddate.value)
        End If
    End If
     
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.text = "N" Then
                
        TXTNoteID = CStr(new_id("Notes", "NoteID", "", True))
                
        rs.AddNew
    ElseIf Me.TxtModFlg.text = "E" Then
        Cn.Execute "delete FixedAssetInstallmentsDetails where FixedAssetInstallmentsid=" & val(Me.TxtFixedAssetInstallmentsid.text)
   
    End If
    
    rs("FixedAssetInstallmentsid").value = val(Me.TxtFixedAssetInstallmentsid.text)
    rs("RecordDate").value = DPRecorddate.value
    rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
    rs("GroupD").value = IIf(Me.dcGroups.BoundText = "", Null, Me.dcGroups.BoundText)
 
    rs("Remarks").value = IIf(Me.txtRemarks.text = "", 0, Me.txtRemarks.text)
     
    rs("Month").value = IIf(Me.CmbMonth.ListIndex = -1, Null, Me.CmbMonth.ListIndex + 1)
    rs("Year").value = IIf(Me.CboYear.ListIndex = -1, Null, val(Me.CboYear.text) + 5)
    rs("NoteId").value = val(TXTNoteID.text)
    rs("NoteSerial") = Me.TxtNoteSerial.text
       
    If ChkForAllAssets.value = vbUnchecked Then
        rs("ForAllAssets").value = 0
    Else
        rs("ForAllAssets").value = 1
    End If

    rs.update
    
    Set RsDev = New ADODB.Recordset
        
    'RsDev.Open "FixedAssetInstallmentsDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
       StrSQL = "SELECT * from dbo.FixedAssetInstallmentsDetails Where (FixedAssetInstallmentsid = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    Dim i As Integer
    Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
    cProgress.StartProgress

    DoEvents
 
    With Me.Grid

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("FixedAssetID")) <> "" Then
         
                RsDev.AddNew
                RsDev("FixedAssetInstallmentsid").value = Me.TxtFixedAssetInstallmentsid.text
                RsDev("FixedAssetID").value = val(.TextMatrix(i, .ColIndex("FixedAssetId")))
                RsDev("CurrentValue").value = .TextMatrix(i, .ColIndex("CurrentValue"))
                RsDev("InstallmentID").value = val(.TextMatrix(i, .ColIndex("InstallmentID")))
                RsDev("InstallmentValue").value = val(.TextMatrix(i, .ColIndex("InstallmentValue")))
                RsDev("InstallmentDate").value = DPRecorddate.value
                RsDev("AccDepreciation").value = val(.TextMatrix(i, .ColIndex("AccDepreciation")))
                RsDev("RemainInstallments").value = val(.TextMatrix(i, .ColIndex("RemainInstallments")))
                RsDev("Month").value = IIf(Me.CmbMonth.ListIndex = -1, Null, Me.CmbMonth.ListIndex + 1)
                RsDev("Year").value = IIf(Me.CboYear.ListIndex = -1, Null, Me.CboYear.ListIndex + 2012)
                RsDev("AddValue").value = val(.TextMatrix(i, .ColIndex("AddValue")))
                RsDev("InstalVal").value = val(.TextMatrix(i, .ColIndex("InstalVal")))
                RsDev("InstallmentProduct").value = 1
                RsDev.update
     
            End If
            
            '
        Next i

    End With

    If CreateJL = False Then '«‰‘«¡ «·ﬁÌÊœ
        GoTo ErrTrap
    End If

    Cn.CommitTrans
    BeginTrans = False
 
    CuurentLogdata

    Select Case Me.TxtModFlg.text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '  Fg_Journal.Enabled = False
    End Select
    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing


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

Private Sub Cmd_Click(Index As Integer)
     On Error GoTo ErrTrap

    Select Case Index

        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            
 
            TxtModFlg.text = "N"
        
            clear_all Me
            ChkForAllAssets.Enabled = True
            Me.TxtFixedAssetInstallmentsid.text = CStr(new_id("FixedAssetInstallments", "FixedAssetInstallmentsid", "", True))
       
            DPRecorddate.value = Date
       
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.rows = 1

            DPRecorddate.value = MonthLastDay(Date)
            Me.dcBranch.BoundText = branch_id
          
        Case 1
    If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            If val(TxtFixedAssetInstallmentsid.text) = 0 Then Exit Sub
            TxtModFlg.text = "E"
            Grid.rows = Grid.rows + 1
            Grid.Enabled = True
            ChkForAllAssets.Enabled = True
            CuurentLogdata

        Case 2
             If Me.CmbMonth.ListIndex = -1 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "«Œ — «·‘Â— «Ê·«"
                        Else
                            MsgBox "    Specify Month"
                        End If

            Exit Sub
        End If

        If Me.CboYear.ListIndex = -1 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "«Œ — «·”‰… «Ê·«"
                    Else
                        MsgBox "    Specify Year"
                    End If

            Exit Sub
        End If
  
  
                      If ChekClodePeriod(DPRecorddate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Because This is Period is Closed"
              End If
              Exit Sub
              End If
If Grid.rows = 1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „ ⁄—÷ «Ì »Ì«‰«  «÷€ÿ ⁄·Ï “— ⁄—÷", vbInformation
Else
MsgBox "No Data Press View", vbInformation
End If
              Exit Sub
End If
          '  CMDView_Click
     If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            
             If ChekClodePeriod(DPRecorddate.value) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
    Else
    MsgBox "Please Change Date Because This is Period is Closed"
    End If
    Exit Sub
    End If
            SaveData
            UpdateFixedAssetPurchaseInformations False
            ChkForAllAssets.Enabled = False
    
        Case 3
            Undo

        Case 4
                    If ChekClodePeriod(DPRecorddate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Because This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(TxtFixedAssetInstallmentsid.text) = 0 Then Exit Sub
            Del_Trans
            UpdateFixedAssetPurchaseInformations True

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 3
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            '   ViewDataList
    
        Case 11
          If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            
            ShowGL_cc Me.TxtNoteSerial.text, , 200, Me.TXTNoteID.text
    End Select

    Exit Sub
ErrTrap:

End Sub

Public Function UpdateFixedAssetPurchaseInformations(delete As Boolean)

    Dim sql As String
    Dim i As Integer
    Dim AccDepreciation As Double
    Dim RemainInstallments As Double
    Dim noOfInstallments As Double
    Dim FixedassetId As Integer
    Dim EXEInstallments   As Double
    Dim currentvalue As Double
    Dim purchaseprice As Double
    Dim KhordaPrice As Double
    Dim cProgress As ClsProgress
    Set cProgress = New ClsProgress
    cProgress.ProgressType = Waiting
    cProgress.StartProgress

    DoEvents

    With Me.Grid

        For i = .FixedRows To .rows - 1
            FixedassetId = val(.TextMatrix(i, .ColIndex("FixedAssetID")))

            If FixedassetId <> 0 Then
        
                GetAllDataAboutFixedAsset FixedassetId, , , , , , , , , , , , , noOfInstallments, , , purchaseprice, , , KhordaPrice
                GetFixedAssetHistory FixedassetId, AccDepreciation, RemainInstallments
                EXEInstallments = (noOfInstallments - RemainInstallments)
                currentvalue = purchaseprice - (KhordaPrice + AccDepreciation)
                
                
                sql = "update FixedAssets set AccDepreciation=" & AccDepreciation & ", EXEInstallments=" & EXEInstallments & ",RemainInstallments=" & IIf(RemainInstallments = -1, 0, RemainInstallments) & ",CurrentValue=" & currentvalue & ",LastDepreciationDate= " & SQLDate(DPRecorddate.value, True)
                sql = sql & "  where id=" & FixedassetId
                Cn.Execute sql
            End If
        
        Next i

    End With

    DoEvents
    cProgress.FinishProgress
    cProgress.StopProgess
    Set cProgress = Nothing
End Function

Private Sub Del_Trans()
    Dim sql As String

    Dim Msg As String
    On Error GoTo ErrTrap

    If TxtFixedAssetInstallmentsid.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtFixedAssetInstallmentsid.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                sql = "Delete   from notes where NoteID=" & val(TXTNoteID.text)
                Cn.Execute sql
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
             
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

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub
  
Private Sub CMDView_Click()
Dim allBranch As Integer
    If Me.TxtModFlg = "R" Then Exit Sub

'If ChkForAllAssets.value = vbChecked Then
'allBranch = 1
'Else
'allBranch = 0
'End If


    If CheckLastInstallmentDate(Me.CmbMonth.ListIndex + 1, val(Me.CboYear.text), val(dcBranch.BoundText)) = True Then
        ViewInstallmentInformations
    Else
        CboYear.ListIndex = -1
        CmbMonth.ListIndex = -1
        Grid.Clear flexClearScrollable, flexClearEverything
        Grid.rows = 1
    End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    'CmdView_Click
    TxtNoteSerial.text = ""
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DCGroups_Click(Area As Integer)
    'CmdView_Click
End Sub

Private Sub DpRecordDate_Change()
    TxtNoteSerial.text = ""
End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    ScreenNameArabic = "«’œ«— «ﬁ”«ÿ «·«Â·«ﬂ"
    ScreenNameEnglish = "Dep Installments"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
    Dim IntDefIndex As Integer
    Dim i As Integer
    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2007 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture

    Dim My_SQL As String

    'If SystemOptions.UserInterface = ArabicInterface Then
    'My_SQL = " select branch_id,branch_name from branches"
    'Else
    'My_SQL = " select branch_id,branch_namee from branches"
    'End If
    'fill_combo Dcbranch, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.dcBranch
 
    My_SQL = " select  GroupID,GroupName from FixedAssetsGroup"
    fill_combo dcGroups, My_SQL
 
    Set BKGrndPic = New ClsBackGroundPic

    With Me.Grid
        .rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
     
        chagelang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From FixedAssetInstallments   where 1=1"
    StrSQL = StrSQL & "  AND  BranchID  is null  or      BranchID in(" & Current_branchSql & ")"
    
     '         If SystemOptions.usertype <> UserAdmin Then
     '   StrSQL = StrSQL & " AND   BranchID=" & Current_branch
    'End If
    
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub
 
Public Sub get_all_employee()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer

    sql = "Select * from emp_all_details "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid

        .rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = 1 To Rs3.RecordCount
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
                       
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
                       
                Rs3.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Sub

Public Sub FillGridWithData()
    Exit Sub

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String

    On Error GoTo ErrTrap
 
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,project_id
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
               
                .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
            
                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
                 "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
           
                '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
                '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
                               
                .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
            
                rs.MoveNext
            
            Next

            rs.Close
        End If

        .rows = .rows + 1
        .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        .IsSubtotal(.rows - 1) = True
        Dim SngTotal As Single
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .rows - 1, .ColIndex("Emp_Salary"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows - 1, .ColIndex("EmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
        net_value = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .rows - 1, .ColIndex("CorrectEmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .rows - 1, .ColIndex("Emp_Salary_sakn"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .rows - 1, .ColIndex("Emp_Salary_bus"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .rows - 1, .ColIndex("Emp_Salary_food"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .rows - 1, .ColIndex("Emp_Salary_others"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .rows - 1, .ColIndex("OverTimePrice"))
        .TextMatrix(.rows - 1, .ColIndex("OverTimePrice")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .rows - 1, .ColIndex("total1"))
        .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .rows - 1, .ColIndex("total2"))
        .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .rows - 1, .ColIndex("Emp_Salary_mang"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .rows - 1, .ColIndex("Emp_Salary_mob"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
    
        .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
        .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
        .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
        .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With

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

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("FixedAssetID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
  
            End If

        Next i

        If .rows > 1 Then
            Me.TxtValue.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("InstallmentValue"), .rows - 1, .ColIndex("InstallmentValue"))
               Me.TxtValue.text = Round(Me.TxtValue.text, 2)
        End If

    End With

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 2
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.TxtFixedAssetInstallmentsid.text = IIf(IsNull(rs("FixedAssetInstallmentsid").value), "", rs("FixedAssetInstallmentsid").value)
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TXTNoteID.text = IIf(IsNull(rs("Noteid").value), "", rs("Noteid").value)
 
    DPRecorddate.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    Me.dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)

    Me.dcGroups.BoundText = IIf(IsNull(rs("GroupD").value), "", rs("GroupD").value)
    Me.txtRemarks.text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
    CmbMonth.ListIndex = IIf(IsNull(rs("Month").value), -1, rs("Month").value - 1)
    CboYear.text = IIf(IsNull(rs("RecordDate").value), year(Date), year(rs("RecordDate").value))

    If IsNull(rs("ForAllAssets").value) Then
        ChkForAllAssets.value = vbUnchecked
    Else

        If rs("ForAllAssets").value = 0 Then
            ChkForAllAssets.value = vbUnchecked
        Else
            ChkForAllAssets.value = vbChecked
        End If
    End If
    If Not (IsNull(rs("LockedInterval").value)) Then
   If rs("LockedInterval").value = True Then
           Cmd(1).Enabled = False
           Cmd(4).Enabled = False
        Else
           Cmd(1).Enabled = True
           Cmd(4).Enabled = True
        End If
     End If
    StrSQL = "select * from FixedAssetInstallmentsDetails where FixedAssetInstallmentsid=" & val(TxtFixedAssetInstallmentsid)
 
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
 
                .TextMatrix(i, .ColIndex("FixedAssetID")) = IIf(IsNull(RsDev("FixedAssetID").value), "", RsDev("FixedAssetID").value)
            
            .TextMatrix(i, .ColIndex("Fullcode")) = getFixedAsstName(val(.TextMatrix(i, .ColIndex("FixedAssetID"))), "Fullcode")
                .TextMatrix(i, .ColIndex("FixedAssetName")) = getFixedAsstName(val(.TextMatrix(i, .ColIndex("FixedAssetID"))), "Name")
                .TextMatrix(i, .ColIndex("CurrentValue")) = IIf(IsNull(RsDev("CurrentValue").value), 0, RsDev("CurrentValue").value)
                .TextMatrix(i, .ColIndex("InstallmentID")) = IIf(IsNull(RsDev("InstallmentID").value), "", RsDev("InstallmentID").value)
                .TextMatrix(i, .ColIndex("InstallmentValue")) = IIf(IsNull(RsDev("InstallmentValue").value), "", RsDev("InstallmentValue").value)
                .TextMatrix(i, .ColIndex("AddValue")) = IIf(IsNull(RsDev("AddValue").value), "", RsDev("AddValue").value)
                .TextMatrix(i, .ColIndex("InstalVal")) = IIf(IsNull(RsDev("InstalVal").value), "", RsDev("InstalVal").value)
                .TextMatrix(i, .ColIndex("InstallmentDate")) = IIf(IsNull(RsDev("InstallmentDate").value), "", RsDev("InstallmentDate").value)
                .TextMatrix(i, .ColIndex("AccDepreciation")) = IIf(IsNull(RsDev("AccDepreciation").value), "", RsDev("AccDepreciation").value)
                .TextMatrix(i, .ColIndex("RemainInstallments")) = IIf(IsNull(RsDev("RemainInstallments").value), "", RsDev("RemainInstallments").value)
            
                RsDev.MoveNext
            Next i
 
        End With

    End If

    Me.TxtModFlg = "R"
    TxtModFlg_Change
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.text = "N" Then
        'CmdRemove.Enabled = True
        'Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.text = "E" Then
        'CmdRemove.Enabled = True
        'Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        'Ele(1).Enabled = False

        'CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
       ' Cmd(1).Enabled = True
      ' Cmd(4).Enabled = True
Cmd(4).Enabled = True
        Cmd(5).Enabled = True

    End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    'On Error GoTo ErrTrap
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

Function chagelang()
    SetInterface Me
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    Cmd(11).Caption = "Print Ge"
    
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(6).Caption = "Exit"

    Me.Cmd(8).Caption = "Issue"
    Me.Cmd(5).Caption = "Cancel"

    Me.Caption = "Fixed Asset Installments"
    Ele(0).Caption = Me.Caption
 
    lbl(4).Caption = "NO"
    ChkForAllAssets.Caption = "For All Assets"
    lbl(17).Caption = "Date"
    lbl(0).Caption = "Remark"
    lbl(15).Caption = "Branch"
    lbl(14).Caption = "Group"
    lbl(3).Caption = "Period"
    
    Ele(3).Caption = "Date"
    lbl(1).Caption = "Month"
    lbl(2).Caption = "Year"
    CMDView.Caption = "View"
    Label1.Caption = "GE No."
    Label3.Caption = "Totals"
    lbl(7).Caption = "Current Rec"
    lbl(6).Caption = "All Rec"
       
    lbl(5).Caption = "User"
       
    With Me.Grid
        .TextMatrix(0, .ColIndex("LineNo")) = "Ser"
        .TextMatrix(0, .ColIndex("Fullcode")) = "F.A Code "
        .TextMatrix(0, .ColIndex("FixedAssetName")) = "F.A  Name "
        .TextMatrix(0, .ColIndex("CurrentValue")) = "Curr. Value "

        .TextMatrix(0, .ColIndex("InstallmentID")) = "Installmen NO"
        .TextMatrix(0, .ColIndex("InstallmentDate")) = "Installmen Date"
        .TextMatrix(0, .ColIndex("AccDepreciation")) = "Acc.Depreciation"
        '.TextMatrix(0, .ColIndex("InstallmenDate")) = " Installment Date"
        .TextMatrix(0, .ColIndex("InstallmentValue")) = "Total Installment Value"
        .TextMatrix(0, .ColIndex("InstalVal")) = "Installment Value"
        .TextMatrix(0, .ColIndex("AddValue")) = "Add Value"
        .TextMatrix(0, .ColIndex("RemainInstallments")) = "Remain Installments"

    End With

End Function


Function GetCarIDByFixedassetid(Optional FixedassetId As Double) As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     id"
sql = sql & " From dbo.TblCarsData"
sql = sql & " Where (fixedassetid = " & FixedassetId & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetCarIDByFixedassetid = IIf(IsNull(rs2("id").value), 0, rs2("id").value)
Else
GetCarIDByFixedassetid = 0
End If
End Function

Function CreateJL() As Boolean
    CreateJL = False
    Dim LngDevID As Long
    Dim DepitAccount As String
    Dim CreditAccount1 As String
    Dim CreditAccount2 As String
    Dim GroupID As Integer
    Dim BranchID As Integer
    Dim FixedassetId As Integer

    Dim Msg As String

    Dim sql As String
    sql = "Delete   from notes where NoteID=" & val(TXTNoteID.text)
    Cn.Execute sql
    '«‰‘«¡ «·ﬁÌÊœ
 
    Dim RsNotes As ADODB.Recordset
    Dim RsDev As ADODB.Recordset
    Dim NoteID As String
    Set RsNotes = New ADODB.Recordset
 '   RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    
    Set RsDev = New ADODB.Recordset
 '   RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  
  
  StrSQL = " SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.* FROM         dbo.DOUBLE_ENTREY_VOUCHERS WHERE     (Double_Entry_Vouchers_ID = - 1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


   StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    RsNotes.AddNew
    
    RsNotes("NoteID").value = CStr(TXTNoteID.text)
    RsNotes("Note_Value").value = val(Me.TxtValue.text)

    If SystemOptions.UserInterface = ArabicInterface Then
        RsNotes("Remark").value = "ﬁÌœ «·«Â·«ﬂ «·‘Â—Ì ··«’Ê· ⁄‰ ‘Â—  " & Me.CmbMonth.ListIndex + 1 & " ·”‰… " & Me.CboYear.text '.ListIndex + 2012
    Else
        RsNotes("Remark").value = "Depreciation Monthly Jl Entry Month:   " & Me.CmbMonth.ListIndex + 1 & " Year " & Me.CboYear.text ' + 2012
    End If

    my_branch = 1
    RsNotes("NoteType").value = 90
    RsNotes("NoteDate").value = DPRecorddate.value
    RsNotes("UserID").value = user_id
    RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '„”·”· «·ﬁÌœ
    RsNotes("NoteSerial1") = TxtFixedAssetInstallmentsid.text
    RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
    RsNotes("sanad_year").value = year(DPRecorddate.value)
    RsNotes("sanad_month").value = Month(DPRecorddate.value)
    RsNotes("branch_no").value = my_branch 'Val(Me.DcBranch.BoundText)
    RsNotes("note_value_by_characters").value = WriteNo(Format(val(TxtValue.text), "0.00"), 0, True, ".")
    RsNotes.update
    Dim des As String
    Dim i As Integer
    Dim lineno As Integer
    lineno = 0

    With Grid

        For i = 1 To .rows - 1

            If val(.TextMatrix(i, .ColIndex("FixedAssetID"))) <> 0 Then
                FixedassetId = val(.TextMatrix(i, .ColIndex("FixedAssetID")))
'                GetAllDataAboutFixedAsset FixedassetId, , GroupID, BranchId
             '   GetFixedAssetsGroupAccount GroupID, 25, BranchId, , , , , , , DepitAccount      'Õ”«» „’—Ê›«  «·«’·
             '   GetFixedAssetsGroupAccount GroupID, 26, BranchId, , , , , , , , CreditAccount1     '„Ã„⁄ «·«Â·«ﬂ

'GetFixedAssetsGroupAccount GroupID, 25, BranchId, , , , , , , DepitAccount      'Õ”«» „’—Ê›«  «·«’·
'                GetFixedAssetsGroupAccount GroupID, 26, BranchId, , , , , , , DepitAccount, CreditAccount1      '„Ã„⁄ «·«Â·«ﬂ


 BranchID = val(.TextMatrix(i, .ColIndex("Branch_NO")))
DepitAccount = (.TextMatrix(i, .ColIndex("DepitAccount")))
CreditAccount1 = (.TextMatrix(i, .ColIndex("CreditAccount")))
Dim CarID As Double

  ' fixedassetid = val(.TextMatrix(LngRow, .ColIndex("FixedAssetID")))
   CarID = GetCarIDByFixedassetid(CDbl(FixedassetId))
   
                If SystemOptions.UserInterface <> ArabicInterface Then
                    des = "  Fixed Asset Installment No " & val(.TextMatrix(i, .ColIndex("InstallmentID"))) & "   Asset Name:  '" & .TextMatrix(i, .ColIndex("FixedAssetName"))
                Else
                    des = "»‰«¡ ⁄·Ï «’œ«— «·ﬁ”ÿ —ﬁ„  " & val(.TextMatrix(i, .ColIndex("InstallmentID"))) & "  ··«’·  '" & .TextMatrix(i, .ColIndex("FixedAssetName"))
                End If
           
                lineno = lineno + 1
                LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)

                '„œÌ‰
                If ModAccounts.AddNewDev(LngDevID, lineno, DepitAccount, val(.TextMatrix(i, .ColIndex("InstallmentValue"))), 0, des, val(Me.TXTNoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.DPRecorddate.value, user_id, , , , , , , , , , , , , , , FixedassetId, GroupID, BranchID, BranchID, CarID) = False Then
                    'GoTo ErrTrap
                    
                End If

                lineno = lineno + 1

                '            œ«∆‰ 1
                If ModAccounts.AddNewDev(LngDevID, lineno, CreditAccount1, val(.TextMatrix(i, .ColIndex("InstallmentValue"))), 1, des, val(Me.TXTNoteID), , , SystemOptions.SysCurrentAccountIntervalID, Me.DPRecorddate.value, user_id, , , , , , , , , , , , , , , FixedassetId, GroupID, BranchID, BranchID) = False Then
                    '  GoTo ErrTrap
                    
                End If
            End If
    
        Next i
        
    End With

    CreateJL = True
    Exit Function
ErrTrap:
    CreateJL = False
End Function

