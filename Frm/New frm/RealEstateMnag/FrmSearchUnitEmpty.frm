VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSerachUnitEmpty 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ «·ÊÕœ«  «·‘«€—Â"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17820
   Icon            =   "FrmSearchUnitEmpty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   17820
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   18720
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   18240
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   18720
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   18840
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   17940
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   -360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   18195
      _cx             =   32094
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
      Caption         =   "«·»ÕÀ ⁄‰ «·ÊÕœ«  «·‘«€—Â"
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
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   6000
         Picture         =   "FrmSearchUnitEmpty.frx":038A
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
         TabIndex        =   23
         Top             =   480
         Width           =   2205
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   1230
      TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         Left            =   3825
         TabIndex        =   7
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
         TabIndex        =   8
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
         Left            =   855
         TabIndex        =   9
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
         TabIndex        =   18
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
         TabIndex        =   27
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   9120
      TabIndex        =   10
      Top             =   10200
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
      Left            =   18000
      TabIndex        =   11
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
      Left            =   18360
      TabIndex        =   20
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   8535
      Left            =   0
      TabIndex        =   28
      Top             =   600
      Width           =   17880
      _cx             =   31538
      _cy             =   15055
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
      Caption         =   "«·»Ì«‰« |Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "FrmSearchUnitEmpty.frx":3FF2
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   8070
         Left            =   18525
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   45
         Width           =   17790
         _cx             =   31380
         _cy             =   14235
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
            TabIndex        =   30
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
            FormatString    =   $"FrmSearchUnitEmpty.frx":438C
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
            TabIndex        =   41
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
            TabIndex        =   31
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8070
         Index           =   15
         Left            =   45
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   45
         Width           =   17790
         _cx             =   31380
         _cy             =   14235
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
         _GridInfo       =   $"FrmSearchUnitEmpty.frx":44D8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8040
            Index           =   16
            Left            =   15
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   15
            Width           =   17760
            _cx             =   31327
            _cy             =   14182
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
            Begin VB.CommandButton Command2 
               Caption         =   "ÿ»«⁄…"
               Height          =   360
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   7605
               Width           =   2055
            End
            Begin VB.CommandButton Command1 
               Caption         =   "„”Õ"
               Height          =   360
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   7605
               Width           =   2055
            End
            Begin VB.CommandButton BtonAdd 
               Caption         =   "»ÕÀ"
               Height          =   360
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   7605
               Width           =   2055
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Height          =   4680
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   2835
               Width           =   17760
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   4365
                  Left            =   0
                  TabIndex        =   54
                  Top             =   240
                  Width           =   17685
                  _cx             =   31194
                  _cy             =   7699
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
                  Cols            =   35
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmSearchUnitEmpty.frx":450E
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
                  AutoSizeMouse   =   0   'False
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
            End
            Begin VB.Frame Frame11 
               Height          =   2880
               Left            =   75
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   0
               Width           =   12825
               Begin VB.TextBox TxtSearch 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3360
                  TabIndex        =   84
                  Top             =   600
                  Width           =   855
               End
               Begin VB.ComboBox DcbValue 
                  Height          =   315
                  ItemData        =   "FrmSearchUnitEmpty.frx":4A38
                  Left            =   120
                  List            =   "FrmSearchUnitEmpty.frx":4A3A
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   2400
                  Width           =   735
               End
               Begin VB.ComboBox DcbLenth 
                  Height          =   315
                  ItemData        =   "FrmSearchUnitEmpty.frx":4A3C
                  Left            =   2640
                  List            =   "FrmSearchUnitEmpty.frx":4A3E
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   1320
                  Width           =   735
               End
               Begin VB.ComboBox DcbhaveFurniture 
                  Height          =   315
                  ItemData        =   "FrmSearchUnitEmpty.frx":4A40
                  Left            =   120
                  List            =   "FrmSearchUnitEmpty.frx":4A4A
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   960
                  Width           =   1575
               End
               Begin VB.TextBox TxtACCount 
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
                  Left            =   2640
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   2400
                  Width           =   1575
               End
               Begin VB.TextBox TxtRentValue 
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
                  Left            =   840
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   2400
                  Width           =   855
               End
               Begin VB.TextBox TxtWC 
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
                  Left            =   2640
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   2040
                  Width           =   1575
               End
               Begin VB.TextBox TxtLoungeCount 
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
                  TabIndex        =   66
                  Top             =   2040
                  Width           =   1575
               End
               Begin VB.TextBox Txtkithchencount 
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
                  TabIndex        =   65
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.TextBox TxtRoom 
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
                  Left            =   2640
                  MaxLength       =   50
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   1680
                  Width           =   1575
               End
               Begin VB.TextBox TxtFloor 
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
                  TabIndex        =   63
                  Top             =   1320
                  Width           =   1575
               End
               Begin VB.TextBox Txtlength 
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
                  TabIndex        =   62
                  Top             =   1320
                  Width           =   855
               End
               Begin VB.ListBox ListGroupSelected 
                  Height          =   2400
                  ItemData        =   "FrmSearchUnitEmpty.frx":4A5E
                  Left            =   5280
                  List            =   "FrmSearchUnitEmpty.frx":4A65
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   240
                  Width           =   3375
               End
               Begin VB.ListBox ListGroupAll 
                  Height          =   2400
                  ItemData        =   "FrmSearchUnitEmpty.frx":4A7C
                  Left            =   9360
                  List            =   "FrmSearchUnitEmpty.frx":4A83
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   240
                  Width           =   3375
               End
               Begin MSDataListLib.DataCombo DcbUnit 
                  Bindings        =   "FrmSearchUnitEmpty.frx":4A95
                  Height          =   315
                  Left            =   2640
                  TabIndex        =   57
                  Top             =   960
                  Width           =   1575
                  _ExtentX        =   2778
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
               Begin MSDataListLib.DataCombo DcboCityID 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   60
                  Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
                  Top             =   240
                  Width           =   4095
                  _ExtentX        =   7223
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo dcbAqarType 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   85
                  Top             =   600
                  Width           =   3255
                  _ExtentX        =   5741
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„  «·⁄ﬁ«—"
                  Height          =   285
                  Index           =   0
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   600
                  Width           =   1050
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ﬁÌ„Â "
                  Height          =   285
                  Index           =   14
                  Left            =   1680
                  TabIndex        =   80
                  Top             =   2400
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄œœ «·’«·« "
                  Height          =   285
                  Index           =   13
                  Left            =   1680
                  TabIndex        =   79
                  Top             =   2040
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Index           =   12
                  Left            =   1440
                  TabIndex        =   78
                  Top             =   1680
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄œœ «·„ﬂÌ›« "
                  Height          =   285
                  Index           =   11
                  Left            =   4080
                  TabIndex        =   77
                  Top             =   2400
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "œÊ—«  «·„Ì«Â"
                  Height          =   285
                  Index           =   10
                  Left            =   4080
                  TabIndex        =   74
                  Top             =   2040
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄œœ «·„ÿ«»Œ"
                  Height          =   285
                  Index           =   9
                  Left            =   1680
                  TabIndex        =   73
                  Top             =   1680
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "⁄œœ «·€—›"
                  Height          =   285
                  Index           =   4
                  Left            =   4320
                  TabIndex        =   72
                  Top             =   1680
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "—ﬁ„ «·ÿ«»ﬁ"
                  Height          =   285
                  Index           =   3
                  Left            =   1680
                  TabIndex        =   71
                  Top             =   1320
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„”«Õ…«·ÊÕœÂ"
                  Height          =   285
                  Index           =   2
                  Left            =   4320
                  TabIndex        =   70
                  Top             =   1320
                  Width           =   885
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„ «·ÕÌ"
                  Height          =   285
                  Index           =   5
                  Left            =   4035
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   240
                  Width           =   1050
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   " √ÀÌÀ «·ÊÕœÂ"
                  Height          =   285
                  Index           =   1
                  Left            =   1680
                  TabIndex        =   59
                  Top             =   960
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "‰Ê⁄ «·ÊÕœÂ"
                  Height          =   285
                  Index           =   5
                  Left            =   4320
                  TabIndex        =   58
                  Top             =   960
                  Width           =   885
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
                  Height          =   255
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   1440
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
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   1080
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
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   720
                  Width           =   495
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
                  Left            =   8760
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   360
                  Width           =   495
               End
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Height          =   3000
               Index           =   11
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   0
               Width           =   17760
               Begin VB.Frame Frame1 
                  Height          =   2895
                  Left            =   12960
                  TabIndex        =   55
                  Top             =   120
                  Width           =   4815
                  Begin VB.Label lblCompanyname 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·”« —Ì…"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   27.75
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00008000&
                     Height          =   5295
                     Left            =   120
                     TabIndex        =   56
                     Top             =   1560
                     Width           =   2895
                  End
                  Begin VB.Image Image1 
                     Height          =   1335
                     Left            =   120
                     Picture         =   "FrmSearchUnitEmpty.frx":4AAA
                     Stretch         =   -1  'True
                     Top             =   120
                     Width           =   4500
                  End
               End
            End
            Begin ImpulseButton.ISButton Accredit 
               Height          =   795
               Left            =   0
               TabIndex        =   40
               Top             =   7995
               Visible         =   0   'False
               Width           =   2580
               _ExtentX        =   4551
               _ExtentY        =   1402
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
               Height          =   660
               Index           =   8
               Left            =   0
               TabIndex        =   51
               Top             =   23910
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   1164
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
               ButtonImage     =   "FrmSearchUnitEmpty.frx":15168
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   615
               Index           =   10
               Left            =   0
               TabIndex        =   52
               Top             =   -5940
               Width           =   960
               _ExtentX        =   1693
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
               ButtonImage     =   "FrmSearchUnitEmpty.frx":15702
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   630
               Index           =   11
               Left            =   -150
               TabIndex        =   53
               Top             =   54180
               Width           =   960
               _ExtentX        =   1693
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
               ButtonImage     =   "FrmSearchUnitEmpty.frx":15C9C
               DrawFocusRectangle=   0   'False
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   8040
            Index           =   9
            Left            =   15
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   15
            Width           =   17760
            _cx             =   31327
            _cy             =   14182
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
               Height          =   6030
               Left            =   4695
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   1590
               Width           =   930
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   4245
               Left            =   5865
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   2205
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   4245
               Index           =   67
               Left            =   3330
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   2205
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   4020
               Index           =   68
               Left            =   5625
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   2595
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
               Height          =   4785
               Index           =   69
               Left            =   4140
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   2205
               Width           =   555
            End
         End
      End
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
      TabIndex        =   26
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
      Left            =   17610
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   11805
      TabIndex        =   17
      Top             =   10275
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2430
      TabIndex        =   16
      Top             =   10350
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   690
      TabIndex        =   15
      Top             =   10350
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   90
      TabIndex        =   14
      Top             =   6900
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1740
      TabIndex        =   13
      Top             =   6900
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   20760
      TabIndex        =   12
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "FrmSerachUnitEmpty"
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
 Dim coun As Integer

Sub retrivetInformationUnites()
  Dim I As Integer
  Dim j As Integer
  Dim k As Integer
 
  Dim Msg As String

  Dim Rs1 As ADODB.Recordset
  Dim sql As String
'ListGroupSelected.Clear

 FG.Clear flexClearScrollable, flexClearEverything
 FG.Rows = 2
 If ListGroupSelected.ListCount > 0 Then
With FG

          For k = 1 To ListGroupSelected.ListCount

    Set Rs1 = New ADODB.Recordset
  sql = "  SELECT  Mremarks, ready,readyDAte,  dbo.TblAqarDetai.id, dbo.TblAqarDetai.unittype, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblAqarDetai.rentType, dbo.TblAqarDetai.meterPrice, dbo.TblAqarDetai.roomscount,"
  sql = sql & "                  dbo.TblAqarDetai.WCcount, dbo.TblAqarDetai.kithchencount, dbo.TblAqarDetai.Id, dbo.TblAqarDetai.length, dbo.TblAqarDetai.unitdesc, dbo.TblAqarDetai.unitno,"
  sql = sql & "                     dbo.TblAqarDetai.RentValue, dbo.TblAqarDetai.customerid, dbo.TblAqarDetai.haveFurniture, dbo.TblAqarDetai.namerentType, dbo.TblAqarDetai.Floor,"
  sql = sql & "                     dbo.TblAqarDetai.LoungeCount, dbo.TblAqarDetai.ACCount, dbo.TblAqarDetai.Services, dbo.TblAqarDetai.Water, dbo.TblAqarDetai.electric, dbo.TblAqarDetai.Status,"
  sql = sql & "                     dbo.TblRentStatus.name AS nameStatus, dbo.TblRentStatus.namee AS nameStatusE, dbo.TblAqarDetai.Aqarid, dbo.TblAqar.aqarname, dbo.TblAqar.BranchId,"
  sql = sql & "                     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAqar.cityid, dbo.TblCountriesGovernments.GovernmentName, dbo.TblAqar.heyid,"
  sql = sql & "                     dbo.TblCountriesGovernmentsCities.CityName , dbo.TblAqar.Aqarid AS AqaridH ,dbo.TblAqarDetai.MiniRentValue "
  sql = sql & "  FROM         dbo.TblRentStatus RIGHT OUTER JOIN"
  sql = sql & "                     dbo.TblCountriesGovernmentsCities RIGHT OUTER JOIN"
  sql = sql & "                     dbo.TblAqar ON dbo.TblCountriesGovernmentsCities.CityID = dbo.TblAqar.heyid LEFT OUTER JOIN"
  sql = sql & "                     dbo.TblCountriesGovernments ON dbo.TblAqar.cityid = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
  sql = sql & "                     dbo.TblBranchesData ON dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id RIGHT OUTER JOIN"
  sql = sql & "                     dbo.TblAqarDetai ON dbo.TblAqar.Aqarid = dbo.TblAqarDetai.Aqarid ON dbo.TblRentStatus.id = dbo.TblAqarDetai.Status LEFT OUTER JOIN"
   sql = sql & "                    dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id"
sql = sql & "  Where (dbo.TblAqarDetai.status = 0) And (dbo.TblAqar.BranchId =" & ListGroupSelected.ItemData(k - 1) & ")"
''///
     If val(dcbAqarType.BoundText) <> 0 And dcbAqarType.Text <> "" Then
        sql = sql & " and dbo.TblAqar.Aqarid =" & val(dcbAqarType.BoundText) & ""
     End If
     If val(DcboCityID.BoundText) <> 0 And DcboCityID.Text <> "" Then
        sql = sql & " and dbo.TblAqar.heyid =" & val(DcboCityID.BoundText) & ""
     End If
     
     If val(DcbUnit.BoundText) <> 0 And DcbUnit.Text <> "" Then
        sql = sql & " and dbo.TblAqarDetai.unittype =" & val(DcbUnit.BoundText) & ""
     End If
      If val(DcbhaveFurniture.ListIndex) <> -1 And DcbhaveFurniture.Text <> "" Then
      
        sql = sql & " and dbo.TblAqarDetai.haveFurniture =" & val(DcbhaveFurniture.ListIndex) & ""
     End If
       If txtLength.Text <> "" Then
     Select Case DcbLenth.ListIndex
        Case 0
          sql = sql & " and dbo.TblAqarDetai.length <" & val(txtLength.Text) & ""
          Case 1
          sql = sql & " and dbo.TblAqarDetai.length >" & val(txtLength.Text) & ""
          Case 2
          sql = sql & " and dbo.TblAqarDetai.length <=" & val(txtLength.Text) & ""
          Case 3
          sql = sql & " and dbo.TblAqarDetai.length >=" & val(txtLength.Text) & ""
          Case 4
           sql = sql & " and dbo.TblAqarDetai.length =" & val(txtLength.Text) & ""
          End Select
        End If
       If TxtFloor.Text <> "" Then
      
        sql = sql & " and dbo.TblAqarDetai.Floor ='" & TxtFloor.Text & "'"
     End If
       If TxtRoom.Text <> "" Then
      
        sql = sql & " and dbo.TblAqarDetai.roomscount =" & val(TxtRoom.Text) & ""
     End If
        If Txtkithchencount.Text <> "" Then
      
        sql = sql & " and dbo.TblAqarDetai.kithchencount =" & val(Txtkithchencount.Text) & ""
     End If
          If TxtWC.Text <> "" Then
      
        sql = sql & " and dbo.TblAqarDetai.WCcount =" & val(TxtWC.Text) & ""
     End If
         If TxtLoungeCount.Text <> "" Then
      
        sql = sql & " and dbo.TblAqarDetai.LoungeCount =" & val(TxtLoungeCount.Text) & ""
     End If
         If TxtAccount.Text <> "" Then
      
        sql = sql & " and dbo.TblAqarDetai.ACCount =" & val(TxtAccount.Text) & ""
     End If
          If TxtRentValue.Text <> "" Then
          Select Case DcbValue.ListIndex
          Case 0
          sql = sql & " and dbo.TblAqarDetai.RentValue <" & val(TxtRentValue.Text) & ""
          Case 1
          sql = sql & " and dbo.TblAqarDetai.RentValue >" & val(TxtRentValue.Text) & ""
          Case 2
          sql = sql & " and dbo.TblAqarDetai.RentValue <=" & val(TxtRentValue.Text) & ""
          Case 3
          sql = sql & " and dbo.TblAqarDetai.RentValue >=" & val(TxtRentValue.Text) & ""
          Case 4
           sql = sql & " and dbo.TblAqarDetai.RentValue =" & val(TxtRentValue.Text) & ""
          End Select
           End If
     
''//
           Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
    If .Rows = 2 Then
   .Rows = .Rows - 1
   Else
   j = .Rows
   End If
   j = .Rows
.Rows = .Rows + Rs1.RecordCount
Rs1.MoveFirst
        For I = j To .Rows - 1
          ' .TextMatrix(i, .ColIndex("Ser")) = i
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(I, .ColIndex("nameStatus")) = IIf(IsNull(Rs1("nameStatus").value), "", Rs1("nameStatus").value)
            .TextMatrix(I, .ColIndex("rentType")) = IIf(IsNull(Rs1("rentType").value), "", Rs1("rentType").value)
             .TextMatrix(I, .ColIndex("name")) = IIf(IsNull(Rs1("name").value), "", Rs1("name").value)
             .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
            Else
               .TextMatrix(I, .ColIndex("name")) = IIf(IsNull(Rs1("namee").value), "", Rs1("namee").value)
               .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
               .TextMatrix(I, .ColIndex("nameStatus")) = IIf(IsNull(Rs1("nameStatusE").value), "", Rs1("nameStatusE").value)
            End If
      .TextMatrix(I, .ColIndex("id")) = IIf(IsNull(Rs1("id").value), "", Rs1("id").value)
       .TextMatrix(I, .ColIndex("ready")) = IIf(IsNull(Rs1("ready").value), 0, Rs1("ready").value)
        .TextMatrix(I, .ColIndex("readyDAte")) = IIf(IsNull(Rs1("readyDAte").value), "", Rs1("readyDAte").value)
        .TextMatrix(I, .ColIndex("Mremarks")) = IIf(IsNull(Rs1("Mremarks").value), "", Rs1("Mremarks").value)
        
      
       .TextMatrix(I, .ColIndex("iqar")) = IIf(IsNull(Rs1("aqarname").value), "", Rs1("aqarname").value)
     .TextMatrix(I, .ColIndex("GovernmentName")) = IIf(IsNull(Rs1("GovernmentName").value), "", Rs1("GovernmentName").value)
     .TextMatrix(I, .ColIndex("CityName")) = IIf(IsNull(Rs1("CityName").value), "", Rs1("CityName").value)
     .TextMatrix(I, .ColIndex("unitno")) = IIf(IsNull(Rs1("unitno").value), "", Rs1("unitno").value)
     .TextMatrix(I, .ColIndex("unittype")) = IIf(IsNull(Rs1("unittype").value), "", Rs1("unittype").value)
      .TextMatrix(I, .ColIndex("Id")) = IIf(IsNull(Rs1("Id").value), "", Rs1("Id").value)
     .TextMatrix(I, .ColIndex("RentValue")) = val(IIf(IsNull(Rs1("RentValue").value), 0, Rs1("RentValue").value))
     .TextMatrix(I, .ColIndex("Floor")) = IIf(IsNull(Rs1("Floor").value), "", Rs1("Floor").value)
      .TextMatrix(I, .ColIndex("namerentType")) = IIf(IsNull(Rs1("namerentType").value), "", Rs1("namerentType").value)
     .TextMatrix(I, .ColIndex("length")) = IIf(IsNull(Rs1("length").value), "", Rs1("length").value)
      .TextMatrix(I, .ColIndex("meterPrice")) = IIf(IsNull(Rs1("meterPrice").value), "", Rs1("meterPrice").value)
     .TextMatrix(I, .ColIndex("BranchId")) = IIf(IsNull(Rs1("BranchId").value), 0, Rs1("BranchId").value)
      .TextMatrix(I, .ColIndex("roomscount")) = IIf(IsNull(Rs1("roomscount").value), "", Rs1("roomscount").value)
     .TextMatrix(I, .ColIndex("LoungeCount")) = IIf(IsNull(Rs1("LoungeCount").value), "", Rs1("LoungeCount").value)
     
     .TextMatrix(I, .ColIndex("MiniRentValue")) = IIf(IsNull(Rs1("MiniRentValue").value), "", Rs1("MiniRentValue").value)
     .TextMatrix(I, .ColIndex("ACCount")) = IIf(IsNull(Rs1("ACCount").value), "", Rs1("ACCount").value)
     .TextMatrix(I, .ColIndex("haveFurniture")) = IIf(IsNull(Rs1("haveFurniture").value), "", Rs1("haveFurniture").value)
     .TextMatrix(I, .ColIndex("kithchencount")) = IIf(IsNull(Rs1("kithchencount").value), "", Rs1("kithchencount").value)
     .TextMatrix(I, .ColIndex("WCcount")) = IIf(IsNull(Rs1("WCcount").value), "", Rs1("WCcount").value)
     .TextMatrix(I, .ColIndex("Statusid")) = IIf(IsNull(Rs1("Status").value), "", Rs1("Status").value)
     .TextMatrix(I, .ColIndex("unitdesc")) = IIf(IsNull(Rs1("unitdesc").value), "", Rs1("unitdesc").value)
     .TextMatrix(I, .ColIndex("Aqarid")) = IIf(IsNull(Rs1("Aqarid").value), "", Rs1("Aqarid").value)
.Row = I
        Rs1.MoveNext
        Next I
End If

       
       
         Next k
'.AutoResize = True

   End With
   Else
   MsgBox "ÌÃ» «Œ Ì«— ›—⁄ «Ê «ﬂÀ—"
   End If
    
    ReLineGrid
End Sub


Private Sub BtonAdd_Click()

retrivetInformationUnites

End Sub

'Private Sub Cmd_Click(Index As Integer)
'
'    ' On Error GoTo ErrTrap
'    Select Case Index
'
'        Case 0
'
'            If DoPremis(Do_New, Me.name, True) = False Then
'                Exit Sub
'            End If
'
'
'
'            TxtModFlg.text = "N"
          

            
         '     GRID2.Clear flexClearScrollable, flexClearEverything
    'GRID2.Rows = 1
'            Me.DCboUserName.BoundText = user_id
          '  TxtPaymentCounts.text = 1
'dcBranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
'
'            Accredit.Enabled = True
'                If SystemOptions.UserInterface = ArabicInterface Then
'                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
'                                                  Else
'                                                    Accredit.Caption = " send to Approval   "
'                                               End If
'
'        Case 1
'
'            If DoPremis(Do_Edit, Me.name, True) = False Then
'                Exit Sub
'            End If
'
'            TxtModFlg.text = "E"
'
'
'            Me.DCboUserName.BoundText = user_id
'
'        Case 2
'
'            Dim Msg As String
'
'            If Trim(dcBranch.BoundText) = "" Then
'                If SystemOptions.UserInterface = EnglishInterface Then
'                    Msg = "Specify Branch"
'                Else
'                    Msg = "Õœœ «·›—⁄ "
'                End If
'
'                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                dcBranch.SetFocus
'                SendKeys "{F4}"
'                Screen.MousePointer = vbDefault
'                Exit Sub
'            End If
'
'            my_branch = Me.dcBranch.BoundText
'
''            SaveData
'
'        Case 3
'            Undo
'
'        Case 4
'
'            If DoPremis(Do_Delete, Me.name, True) = False Then
'                Exit Sub
'            End If
'
'            Del_Trans
'
'        Case 5
'            Load FrmLinkIteminStoreSearch
'            FrmLinkIteminStoreSearch.show vbModal
'
'        Case 6
'            Unload Me
'
''        Case 7
 '           ShowGL_cc Me.TxtNoteSerial.text, , 200
'
'        Case 12
'         RemoveGridRow
'            Case 13
'             Fg.Clear flexClearScrollable, flexClearEverything
' Fg.Rows = 2
'            coun = 0
'                 Case 9
'
'            If DoPremis(Do_Print, Me.name, True) = False Then
'                Exit Sub
'            End If
'
'            If val(Me.XPTxtID.text) <> 0 Then
'                print_report val(Me.XPTxtID.text)
'
'
'            End If
'
'    End Select
'
'    Exit Sub
'ErrTrap:
'End Sub
'Function print_report(Optional NoteSerial As String)
    
     
'    Dim MySQL As String
'    Dim RsData As New ADODB.Recordset
'    Dim xApp As New CRAXDRT.Application
'    Dim xReport As CRAXDRT.Report
'    Dim CViewer As ClsReportViewer
'    Dim StrReportTitle As String
'    Dim StrFileName As String
'    Dim Msg As String
'
'MySQL = " SELECT     dbo.TblLink_Item_To_StoreH.Ind, dbo.TblLink_Item_To_StoreH.LinkType, dbo.TblLink_Item_To_StoreH.UserID, dbo.TblUsers.UserName,"
'MySQL = MySQL & "                      dbo.TblLink_Item_To_StoreH.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblLink_Item_To_StoreH.RecordeDate,"
'MySQL = MySQL & "                       dbo.TblLink_Item_To_StoreH.Remarks, dbo.TblLink_Item_To_StoreH.Selected, dbo.TblLink_Item_To_StoreH.Posted, dbo.TblLink_Item_To_Store_Details2.Ind AS Ind2,"
'MySQL = MySQL & "                       dbo.TblLink_Item_To_Store_Details2.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblLink_Item_To_Store_Details2.ItemID,"
'MySQL = MySQL & "                       dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblLink_Item_To_Store_Details2.LinkType AS linktype2, dbo.TblLink_Item_To_Store_Details2.GroupID,"
'MySQL = MySQL & "                       dbo.Groups.GroupName , dbo.Groups.GroupNamee"
'MySQL = MySQL & "  , dbo.TblItems.ItemCode  FROM         dbo.TblLink_Item_To_StoreH LEFT OUTER JOIN"
'MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.TblLink_Item_To_StoreH.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
'MySQL = MySQL & "                       dbo.TblUsers ON dbo.TblLink_Item_To_StoreH.UserID = dbo.TblUsers.UserID RIGHT OUTER JOIN"
'MySQL = MySQL & "                       dbo.Groups RIGHT OUTER JOIN"
'MySQL = MySQL & "                       dbo.TblLink_Item_To_Store_Details2 ON dbo.Groups.GroupID = dbo.TblLink_Item_To_Store_Details2.GroupID LEFT OUTER JOIN"
'MySQL = MySQL & "                       dbo.TblItems ON dbo.TblLink_Item_To_Store_Details2.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
'MySQL = MySQL & "                       dbo.TblStore ON dbo.TblLink_Item_To_Store_Details2.StoreID = dbo.TblStore.StoreID ON"
'MySQL = MySQL & "                       dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details2.Ind"
'MySQL = MySQL & " Where (dbo.TblLink_Item_To_StoreH.Ind =" & val(XPTxtID.text) & ")"
''MySQL = MySQL & "Where (dbo.TblLink_Item_To_StoreH.Ind = " & val(XPTxtID.text) & ")"
'
'
' 'MySQL = MySQL & "   Where (dbo.TblTreatment.id =" & val(XPTxtID.text) & ")"

' If SystemOptions.UserInterface = ArabicInterface Then
'          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepLinkingIteminStore.rpt"
'     Else
'        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepLinkingIteminStore.rpt"
'       End If
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
        'GetMsgs 138, vbExclamation
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Function
 '   End If
''
 '   Screen.MousePointer = vbArrowHourglass
 '   Set xReport = xApp.OpenReport(StrFileName)
 '   xReport.Database.SetDataSource RsData
'
'    Dim cCompanyInfo As New ClsCompanyInfo
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
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
'      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
'        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
'        StrReportTitle = ""
'        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
'        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
'        'End If
'    End If
'
'    xReport.ParameterFields(3).AddCurrentValue user_name
'       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
''        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
'      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
'' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
''  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
' '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
'
''    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
'    xReport.reporttitle = StrReportTitle
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.Title
'    xReport.ReportAuthor = App.Title
'    Set CViewer = New ClsReportViewer
 '   CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
''
 '   RsData.Close
 '   Set RsData = Nothing
 '   Screen.MousePointer = vbDefault
'
'
 
  
 
'End Function

'Private Sub CmdHelp_Click()
'    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
'    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
'End Sub



Private Sub Command1_Click()
ListGroupSelected.Clear
 FG.Clear flexClearScrollable, flexClearEverything
 FG.Rows = 2
  clear_all Me
End Sub



Private Sub Command2_Click()
GetData
End Sub

Private Sub dcbAqarType_Change()
dcbAqarType_Click (0)
End Sub

Private Sub dcbAqarType_Click(Area As Integer)
      If val(dcbAqarType.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 
    GetIqarCode , , dcbAqarType.BoundText, EmpCode
    
    Me.TxtSearch.Text = EmpCode
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
'If Me.CBTybe.ListIndex = 2 Then
    MySQL = " select * from TblSearchUnitEmpty"

      If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSearchUnitEmpty.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSearchUnitEmpty.rpt"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
  
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Sub GetData()
Dim RsDetails As ADODB.Recordset
Dim I As Integer
Dim STRSQL As String
STRSQL = "Delete From TblSearchUnitEmpty Where 1<>-1"
                Cn.Execute STRSQL, , adExecuteNoRecords
                
        Set RsDetails = New ADODB.Recordset
        RsDetails.Open "TblSearchUnitEmpty", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
With FG

        For I = .FixedRows To .Rows - 1
            If .TextMatrix(I, .ColIndex("branch_name")) <> "" Then
                 RsDetails.AddNew
                    RsDetails("BranchId").value = val(.TextMatrix(I, .ColIndex("BranchId")))
                    RsDetails("branch_name").value = .TextMatrix(I, .ColIndex("branch_name"))
                    RsDetails("GovernmentName").value = .TextMatrix(I, .ColIndex("GovernmentName"))
                    RsDetails("CityName").value = .TextMatrix(I, .ColIndex("CityName"))
                    RsDetails("iqar").value = .TextMatrix(I, .ColIndex("iqar"))
                    RsDetails("unittype").value = val(.TextMatrix(I, .ColIndex("unittype")))
                    RsDetails("Aqarid").value = val(.TextMatrix(I, .ColIndex("Aqarid")))
                    RsDetails("name").value = .TextMatrix(I, .ColIndex("name"))
                    RsDetails("unitno").value = .TextMatrix(I, .ColIndex("unitno"))
                    RsDetails("namerentType").value = .TextMatrix(I, .ColIndex("namerentType"))
                    RsDetails("Floor").value = val(.TextMatrix(I, .ColIndex("Floor")))
                    RsDetails("rentType").value = val(.TextMatrix(I, .ColIndex("rentType")))
                    RsDetails("length").value = val(.TextMatrix(I, .ColIndex("length")))
                    RsDetails("RentValue").value = val(.TextMatrix(I, .ColIndex("RentValue")))
                    RsDetails("meterPrice").value = val(.TextMatrix(I, .ColIndex("meterPrice")))
                    RsDetails("roomscount").value = val(.TextMatrix(I, .ColIndex("roomscount")))
                    RsDetails("unitdesc").value = .TextMatrix(I, .ColIndex("unitdesc"))
                    RsDetails("LoungeCount").value = val(.TextMatrix(I, .ColIndex("LoungeCount")))
                    RsDetails("ACCount").value = val(.TextMatrix(I, .ColIndex("ACCount")))
                    RsDetails("kithchencount").value = val(.TextMatrix(I, .ColIndex("kithchencount")))
                    If .Cell(flexcpChecked, I, .ColIndex("haveFurniture")) = flexChecked Then
                    RsDetails("haveFurniture").value = "„ƒÀÀ"
                    Else
                    RsDetails("haveFurniture").value = "€Ì— „ƒÀÀ"
                    End If
                    RsDetails("WCcount").value = val(.TextMatrix(I, .ColIndex("WCcount")))
                    RsDetails("MiniRentValue").value = val(.TextMatrix(I, .ColIndex("MiniRentValue")))
                 RsDetails.update

            End If

        Next I
  End With
  print_report
End Sub

Private Sub Fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
           With FG
           
           
                If .Cell(flexcpChecked, .Row, .ColIndex("ready")) = flexUnchecked Then
                         .TextMatrix(.Row, .ColIndex("readyDAte")) = ""
                        
           End If
           
        
                   If .Cell(flexcpChecked, .Row, .ColIndex("ready")) = flexChecked And .TextMatrix(.Row, .ColIndex("readyDAte")) = "" Then
                         .TextMatrix(.Row, .ColIndex("readyDAte")) = Date
                        
           End If
           
           End With
End Sub

Private Sub fg_Click()
On Error Resume Next
 

  With FG

        Select Case .Col

            Case 2
            
                If .TextMatrix(.Row, .ColIndex("ready")) = True Then
                           If .TextMatrix(.Row, .ColIndex("readyDAte")) = "" Then
                                         MsgBox "  «œŒ·  «—ÌŒ «· ÃÂÌ“", vbInformation
                                         Exit Sub
                        End If
           End If
'Mremarks
Dim STRSQL As String
             If .TextMatrix(.Row, .ColIndex("ready")) = True Then
             STRSQL = "update TblAqarDetai  set ready=1 ,readyDAte =' " & .TextMatrix(.Row, .ColIndex("readyDAte")) & "'"
             STRSQL = STRSQL & ", Mremarks='" & .TextMatrix(.Row, .ColIndex("Mremarks")) & "'"
          STRSQL = STRSQL & " where id=" & val(.TextMatrix(.Row, .ColIndex("id")))
             
 Cn.Execute STRSQL
 
           MsgBox "«·ÊÕœ… Ã«Â“…", vbInformation
           Else
         STRSQL = "update TblAqarDetai  set ready=0 ,readyDAte =''  "
                      STRSQL = STRSQL & ", Mremarks='" & .TextMatrix(.Row, .ColIndex("Mremarks")) & "'"
                      STRSQL = STRSQL & " where id=" & val(.TextMatrix(.Row, .ColIndex("id")))
                      
 Cn.Execute STRSQL
           MsgBox "«·ÊÕœ…  €Ì— Ã«Â“…", vbCritical
           End If
            End Select
     End With
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
FG.ColComboList(FG.ColIndex("id")) = "..."

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Label5_Click()

    If ListGroupSelected.ListIndex > -1 Then
        ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
    End If

End Sub
Private Sub Label6_Click()
    ListGroupSelected.Clear
End Sub
Private Sub Label7_Click()
    Dim I As Integer
   
    ListGroupSelected.Clear

    For I = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(I)
        ListGroupSelected.ItemData(I) = ListGroupAll.ItemData(I)
    Next I

End Sub
Private Sub Label8_Click()

 If ListGroupAll.ListIndex > -1 Then
    ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
             
    ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)
            End If
        
End Sub










Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim STRSQL As String
    Dim GrdBack As ClsBackGroundPic
 'Dim count As Integer
 coun = 0
    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
  
  

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
    
'    AddTip

DcbLenth.AddItem "<"
DcbLenth.AddItem ">"
DcbLenth.AddItem "<="
DcbLenth.AddItem ">="
DcbLenth.AddItem "="

DcbValue.AddItem "<"
DcbValue.AddItem ">"
DcbValue.AddItem "<="
DcbValue.AddItem ">="
DcbValue.AddItem "="

    FillMylist
    Set Dcombos = New ClsDataCombos
     ' Dcombos.GetUsers Me.DCboUserName
      Dcombos.get«hay Me.DcboCityID
Dcombos.getAkarUnit Me.DcbUnit
Dcombos.GetIqar dcbAqarType
      ListGroupSelected.Clear
 FG.Clear flexClearScrollable, flexClearEverything
 FG.Rows = 2
ErrTrap:
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



'Private Sub TxtModFlg_Change()



''    Dim i As Integer
''    Dim StrSQL As String
''    ListGroupSelected.Clear
''
''Fg.Clear flexClearScrollable, flexClearEverything
''            Fg.Rows = 2
''            Fg.Enabled = True
''    'On Error GoTo ErrTrap
''    If rs.RecordCount < 1 Then



'
'   RsDetails1.MoveNext
'
'  Next i
'  '''''''''''''''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'
   


'Else
'  Dim RsTest As New ADODB.Recordset



'
'    If KeyCode = vbKeyF12 Then
'        If Cmd(0).Enabled = False Then Exit Sub
'        Cmd_Click (0)
'    End If
'
'    If KeyCode = vbKeyF11 Then
'        If Cmd(1).Enabled = False Then Exit Sub
'        Cmd_Click (1)
'    End If
'
'    If KeyCode = vbKeyF10 Then
''        If Cmd(2).Enabled = False Then Exit Sub
 '       Cmd_Click (2)
 '   End If
'
'    If KeyCode = vbKeyF9 Then
'        If Cmd(3).Enabled = False Then Exit Sub
'        Cmd_Click (3)
'    End If
'
'    If KeyCode = vbKeyF8 Then
'        If Cmd(4).Enabled = False Then Exit Sub
'        Cmd_Click (4)
'    End If
'
'    If Shift = 2 Then
'        If KeyCode = vbKeyX Then
'            If Cmd(6).Enabled = False Then Exit Sub
'            Cmd_Click (6)
'        End If
'    End If
'
'    Exit Sub
'ErrTrap:
'End Sub
Private Sub ReLineGrid()
    Dim I As Integer
    Dim IntCounter  As Integer
    
    IntCounter = 0

    With FG

        For I = .FixedRows To .Rows - 1

            If .TextMatrix(I, .ColIndex("branch_name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("Ser")) = IntCounter
                 If val(.TextMatrix(I, .ColIndex("rentType"))) = 1 Then
                .TextMatrix(I, .ColIndex("RentValue")) = val(.TextMatrix(I, .ColIndex("length"))) * val(.TextMatrix(I, .ColIndex("meterPrice")))
                
 
                End If
           
    
        End If
                

        Next I
 
    End With

End Sub
Function FillMylist()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim I As Integer
    sql = " SELECT * from  TblBranchesData"
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListGroupAll.Clear
 

    If rs.RecordCount > 0 Then

        For I = 1 To rs.RecordCount
             
            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupAll.AddItem IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
            Else
                ListGroupAll.AddItem IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
            End If

            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("branch_id").value
            rs.MoveNext
        Next I

    End If

    rs.Close

    'fil


End Function

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        dcbAqarType.BoundText = EmpID
        dcbAqarType_Click (0)
    End If
End Sub
