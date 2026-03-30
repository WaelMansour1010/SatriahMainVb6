VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form formvocatinl 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·» «Ã«“…"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12570
   Icon            =   "formvocationl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9735
   ScaleWidth      =   12570
   Begin VB.CommandButton cmdApi 
      Caption         =   "Load From Web"
      Height          =   450
      Left            =   1710
      RightToLeft     =   -1  'True
      TabIndex        =   206
      Top             =   1185
      Width           =   1935
   End
   Begin VB.TextBox TxtSearchCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   13920
      TabIndex        =   36
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   13440
      TabIndex        =   35
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   13920
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   14040
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   13140
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   7020
      TabIndex        =   7
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   241041409
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DcboEmpName 
      Height          =   315
      Left            =   3900
      TabIndex        =   8
      Top             =   1185
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   1560
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9120
      Width           =   10665
      _cx             =   18812
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
         Left            =   8430
         TabIndex        =   10
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
         Left            =   7575
         TabIndex        =   11
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
         Left            =   6735
         TabIndex        =   12
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
         Left            =   5880
         TabIndex        =   13
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
         Left            =   5025
         TabIndex        =   14
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
         TabIndex        =   15
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
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   375
         Left            =   855
         TabIndex        =   16
         Top             =   75
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
         Left            =   3960
         TabIndex        =   27
         Top             =   75
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
         Left            =   3120
         TabIndex        =   39
         Top             =   75
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   11
         Left            =   1920
         TabIndex        =   205
         Top             =   75
         Width           =   1005
         _ExtentX        =   1773
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
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   312
      Left            =   6360
      TabIndex        =   17
      Top             =   8760
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
      Left            =   13200
      TabIndex        =   18
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
      Left            =   13560
      TabIndex        =   29
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
   Begin MSDataListLib.DataCombo Dcbranch 
      Bindings        =   "formvocationl.frx":038A
      Height          =   315
      Left            =   1440
      TabIndex        =   31
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
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
      Height          =   6975
      Left            =   0
      TabIndex        =   40
      Top             =   1680
      Width           =   12600
      _cx             =   22225
      _cy             =   12303
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
      Caption         =   "«·»Ì«‰«  «·«”«”Ì…|Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "formvocationl.frx":039F
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6510
         Left            =   13245
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   45
         Width           =   12510
         _cx             =   22066
         _cy             =   11483
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
            TabIndex        =   42
            Tag             =   "1"
            Top             =   240
            Width           =   12270
            _cx             =   21643
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
            FormatString    =   $"formvocationl.frx":0739
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
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   4080
            Width           =   3375
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   13080
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6510
         Index           =   15
         Left            =   45
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   45
         Width           =   12510
         _cx             =   22066
         _cy             =   11483
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
         _GridInfo       =   $"formvocationl.frx":087C
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6480
            Index           =   16
            Left            =   15
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   15
            Width           =   12480
            _cx             =   22013
            _cy             =   11430
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
            Begin VB.TextBox TxtDiscouDay 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   5400
               TabIndex        =   203
               TabStop         =   0   'False
               Top             =   3720
               Width           =   1095
            End
            Begin VB.TextBox TxtToalAbsent 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   202
               Top             =   4440
               Width           =   1095
            End
            Begin VB.TextBox TxtAddDay 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   7800
               TabIndex        =   190
               TabStop         =   0   'False
               Top             =   3720
               Width           =   1095
            End
            Begin VB.TextBox TxtDuVocation 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   4440
               Width           =   1095
            End
            Begin VB.TextBox TxtTotalDay 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFC0&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   4440
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox TxtContDay 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   172
               Top             =   3720
               Width           =   1095
            End
            Begin VB.TextBox TxtWithOutSala1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   4440
               Width           =   1095
            End
            Begin VB.TextBox TxtNewAbsent 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   170
               Top             =   4440
               Width           =   1215
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· √‘Ì—…"
               Height          =   1935
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   0
               Width           =   3852
               Begin VB.CheckBox ChkOutAndBack 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Œ—ÊÃ Ê⁄Êœ…"
                  Height          =   192
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   1560
                  Width           =   1452
               End
               Begin VB.CheckBox chkVistCostOnCompany 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄·Ï Õ”«» «·‘—ﬂ…"
                  Height          =   192
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.CheckBox chkVistCostOnEmployee 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄·Ï Õ”«» «·„ÊŸ›"
                  Height          =   192
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.TextBox txtVisaCost 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   240
                  TabIndex        =   125
                  TabStop         =   0   'False
                  Top             =   600
                  Width           =   2412
               End
               Begin VB.TextBox Text2 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   240
                  Width           =   2412
               End
               Begin VB.CheckBox chkOutOnly 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Œ—ÊÃ ›ﬁÿ"
                  Height          =   192
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   1560
                  Width           =   1695
               End
               Begin VB.CheckBox chkForFamily 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "·√›—«œ «·⁄«∆·…"
                  Height          =   192
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   1320
                  Width           =   1695
               End
               Begin VB.CheckBox chkForEmployee 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "··„ÊŸ›"
                  Height          =   192
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   1320
                  Width           =   1695
               End
               Begin MSComCtl2.DTPicker DTPicker1 
                  Height          =   315
                  Left            =   7320
                  TabIndex        =   97
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   240975873
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DataCombo3 
                  Height          =   315
                  Left            =   6840
                  TabIndex        =   98
                  Top             =   1080
                  Width           =   1995
                  _ExtentX        =   3519
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ﬂ«·Ì› «· √‘Ì—…"
                  Height          =   285
                  Index           =   9
                  Left            =   2640
                  TabIndex        =   124
                  Top             =   600
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· √‘Ì—…"
                  Height          =   288
                  Index           =   18
                  Left            =   2676
                  TabIndex        =   108
                  Top             =   240
                  Width           =   1008
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„— »…"
                  Height          =   285
                  Index           =   20
                  Left            =   8520
                  TabIndex        =   105
                  Top             =   1080
                  Width           =   1125
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊŸÌ›…"
                  Height          =   285
                  Left            =   5280
                  TabIndex        =   104
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   19
                  Left            =   6840
                  TabIndex        =   103
                  Top             =   360
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
                  Height          =   285
                  Index           =   17
                  Left            =   8640
                  TabIndex        =   102
                  Top             =   360
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—« » «·«”«”Ì"
                  Height          =   285
                  Index           =   16
                  Left            =   8280
                  TabIndex        =   101
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ﬂ«·Ì› «· √‘Ì—…"
                  Height          =   288
                  Index           =   12
                  Left            =   5040
                  TabIndex        =   100
                  Top             =   720
                  Width           =   1008
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„œÌ—"
                  Height          =   288
                  Index           =   11
                  Left            =   4680
                  TabIndex        =   99
                  Top             =   1080
                  Width           =   1008
               End
            End
            Begin VB.TextBox txtreson 
               Alignment       =   1  'Right Justify
               Height          =   495
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   83
               Top             =   4920
               Width           =   11055
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Height          =   1815
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   1800
               Width           =   3855
               Begin VB.TextBox TxtNoVacation 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1740
                  TabIndex        =   130
                  TabStop         =   0   'False
                  Top             =   1440
                  Width           =   915
               End
               Begin MSComCtl2.DTPicker xpdtbfrom 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   71
                  Top             =   240
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   241041409
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker xpdtbto 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   72
                  Top             =   720
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   241041409
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal fromdateH 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   89
                  Top             =   240
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin Dynamic_Byte.NourHijriCal todateH 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   90
                  Top             =   720
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin Dynamic_Byte.NourHijriCal dtpResumeWorkh 
                  Height          =   312
                  Left            =   120
                  TabIndex        =   122
                  Top             =   1080
                  Width           =   1092
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin MSComCtl2.DTPicker dtpResumeWork 
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   123
                  Top             =   1080
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   241041409
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   285
                  Index           =   37
                  Left            =   120
                  TabIndex        =   181
                  Top             =   1440
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «Ì«„ «·«Ã«“…"
                  Height          =   285
                  Index           =   34
                  Left            =   2640
                  TabIndex        =   131
                  Top             =   1440
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "  «·„»«‘—… «·„ Êﬁ⁄…"
                  Height          =   405
                  Index           =   32
                  Left            =   2640
                  TabIndex        =   75
                  Top             =   960
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«‰ Â«¡ «·«Ã«“…"
                  Height          =   285
                  Index           =   33
                  Left            =   2640
                  TabIndex        =   74
                  Top             =   720
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "»œ«Ì… «·«Ã«“…"
                  Height          =   285
                  Index           =   35
                  Left            =   2640
                  TabIndex        =   73
                  Top             =   240
                  Width           =   1125
               End
            End
            Begin VB.Frame lbltype 
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·«Ã«“…"
               Height          =   735
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   2160
               Width           =   8535
               Begin VB.TextBox Txtother 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   240
                  Width           =   1575
               End
               Begin XtremeSuiteControls.RadioButton Rdb1 
                  Height          =   495
                  Index           =   71
                  Left            =   5880
                  TabIndex        =   67
                  Top             =   120
                  Width           =   1095
                  _Version        =   786432
                  _ExtentX        =   1931
                  _ExtentY        =   873
                  _StockProps     =   79
                  Caption         =   "—”„Ì…"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton Rdb2 
                  Height          =   495
                  Index           =   72
                  Left            =   3960
                  TabIndex        =   68
                  Top             =   120
                  Width           =   1095
                  _Version        =   786432
                  _ExtentX        =   1931
                  _ExtentY        =   873
                  _StockProps     =   79
                  Caption         =   "«÷ÿ—«—Ì…"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton Rdbother 
                  Height          =   495
                  Index           =   73
                  Left            =   1920
                  TabIndex        =   69
                  Top             =   120
                  Width           =   1095
                  _Version        =   786432
                  _ExtentX        =   1931
                  _ExtentY        =   873
                  _StockProps     =   79
                  Caption         =   "«Œ—Ï"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin VB.Frame lblde 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·„ÊŸ›"
               Height          =   2175
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   0
               Width           =   8655
               Begin VB.CheckBox chkWithoutSalary 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Ã«“… »œÊ‰ —« »"
                  Height          =   312
                  Left            =   3360
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   1320
                  Width           =   1692
               End
               Begin VB.CheckBox chkWithSalary 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Ã«“… »—« »"
                  Height          =   312
                  Left            =   6000
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   1320
                  Width           =   1692
               End
               Begin VB.CheckBox chkManagerApprove 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Ê«›ﬁÂ «·„œÌ—"
                  Height          =   312
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   1320
                  Width           =   1692
               End
               Begin MSDataListLib.DataCombo DcboEmpDepartments 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   55
                  Top             =   240
                  Width           =   2115
                  _ExtentX        =   3731
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker DBIssueDate 
                  Height          =   315
                  Left            =   9360
                  TabIndex        =   56
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   243662849
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcboJobsType 
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   57
                  Top             =   240
                  Width           =   4095
                  _ExtentX        =   7223
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboSpecifications 
                  Height          =   315
                  Left            =   8880
                  TabIndex        =   58
                  Top             =   1080
                  Width           =   1995
                  _ExtentX        =   3519
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcmbManagerID 
                  Bindings        =   "formvocationl.frx":08B0
                  Height          =   315
                  Left            =   120
                  TabIndex        =   93
                  Top             =   960
                  Width           =   7575
                  _ExtentX        =   13361
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
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
               Begin MSDataListLib.DataCombo DcbDetpartment 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   135
                  Top             =   600
                  Width           =   7575
                  _ExtentX        =   13361
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin Dynamic_Byte.NourHijriCal lastHolidaydateH 
                  Height          =   315
                  Left            =   420
                  TabIndex        =   136
                  Top             =   1680
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin MSComCtl2.DTPicker lastHolidaydate 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   137
                  Top             =   1680
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   243597313
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal BignDateH 
                  Height          =   315
                  Left            =   4740
                  TabIndex        =   140
                  Top             =   1680
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   556
               End
               Begin MSComCtl2.DTPicker BignDate 
                  Height          =   315
                  Left            =   5880
                  TabIndex        =   141
                  Top             =   1680
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   243597313
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ »œ«Ì… «·⁄„·"
                  Height          =   285
                  Index           =   36
                  Left            =   6900
                  TabIndex        =   142
                  Top             =   1680
                  Width           =   1635
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0C0C0&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ «Œ— „»«‘—…/«Ã«“…"
                  Height          =   285
                  Index           =   39
                  Left            =   3060
                  TabIndex        =   138
                  Top             =   1680
                  Width           =   1635
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«œ«—…"
                  Height          =   285
                  Left            =   7920
                  TabIndex        =   134
                  Top             =   600
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„œÌ—"
                  Height          =   285
                  Index           =   10
                  Left            =   7560
                  TabIndex        =   92
                  Top             =   960
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—« » «·«”«”Ì"
                  Height          =   285
                  Index           =   5
                  Left            =   11520
                  TabIndex        =   64
                  Top             =   360
                  Width           =   1125
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
                  Height          =   285
                  Index           =   13
                  Left            =   8640
                  TabIndex        =   63
                  Top             =   360
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Êﬁ⁄ «·⁄„·"
                  Height          =   285
                  Index           =   15
                  Left            =   2160
                  TabIndex        =   62
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   285
                  Index           =   23
                  Left            =   10080
                  TabIndex        =   61
                  Top             =   360
                  Width           =   885
               End
               Begin VB.Label lblj 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÊŸÌ›…"
                  Height          =   285
                  Left            =   7920
                  TabIndex        =   60
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„— »…"
                  Height          =   285
                  Index           =   14
                  Left            =   8520
                  TabIndex        =   59
                  Top             =   1080
                  Width           =   1125
               End
            End
            Begin XtremeSuiteControls.GroupBox gb 
               Height          =   975
               Left            =   0
               TabIndex        =   76
               Top             =   5400
               Width           =   12495
               _Version        =   786432
               _ExtentX        =   22040
               _ExtentY        =   1720
               _StockProps     =   79
               Caption         =   "Ê”«∆· «· Ê«’· «À‰«¡ «·«Ã«“…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
               Begin VB.TextBox TxtAdress 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   4440
                  TabIndex        =   85
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   7215
               End
               Begin VB.TextBox xptxtphone 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   240
                  TabIndex        =   79
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   3135
               End
               Begin VB.TextBox xptxttelephone 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   0
                  TabIndex        =   78
                  TabStop         =   0   'False
                  Top             =   2160
                  Width           =   3135
               End
               Begin VB.TextBox xptxtother 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   4440
                  TabIndex        =   77
                  TabStop         =   0   'False
                  Top             =   600
                  Width           =   7215
               End
               Begin VB.Label lbladres 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄‰Ê«‰"
                  Height          =   285
                  Left            =   11400
                  TabIndex        =   86
                  Top             =   240
                  Width           =   885
               End
               Begin VB.Label lblmo 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÃÊ«·"
                  Height          =   285
                  Left            =   3120
                  TabIndex        =   82
                  Top             =   240
                  Width           =   885
               End
               Begin VB.Label lbltel 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "À«» "
                  Height          =   285
                  Left            =   3120
                  TabIndex        =   81
                  Top             =   600
                  Width           =   885
               End
               Begin VB.Label lblother 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ—Ï"
                  Height          =   285
                  Left            =   11400
                  TabIndex        =   80
                  Top             =   600
                  Width           =   885
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   735
               Index           =   1
               Left            =   9720
               TabIndex        =   143
               TabStop         =   0   'False
               Top             =   2880
               Width           =   2745
               _cx             =   4842
               _cy             =   1296
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
               Begin VB.TextBox TxtAbsent 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Height          =   285
                  Left            =   -585
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   750
               End
               Begin VB.TextBox TxtDayAbs 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1305
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   360
                  Width           =   600
               End
               Begin VB.TextBox TxtYearAbs 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   150
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   360
                  Width           =   600
               End
               Begin VB.TextBox TxtMoAbs 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   360
                  Width           =   600
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   315
                  Index           =   47
                  Left            =   720
                  TabIndex        =   151
                  Top             =   120
                  Width           =   495
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”‰…"
                  Height          =   315
                  Index           =   46
                  Left            =   150
                  TabIndex        =   150
                  Top             =   120
                  Width           =   480
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÌÊ„"
                  Height          =   315
                  Index           =   45
                  Left            =   1455
                  TabIndex        =   149
                  Top             =   120
                  Width           =   345
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ì«„ «·€Ì«»"
                  ForeColor       =   &H00C00000&
                  Height          =   555
                  Index           =   58
                  Left            =   2175
                  TabIndex        =   148
                  Top             =   0
                  Width           =   495
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   735
               Index           =   2
               Left            =   6960
               TabIndex        =   152
               TabStop         =   0   'False
               Top             =   2880
               Width           =   2745
               _cx             =   4842
               _cy             =   1296
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
               Begin VB.TextBox TxtYear 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   120
                  TabIndex        =   155
                  Top             =   360
                  Width           =   615
               End
               Begin VB.TextBox TxtMonth 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   840
                  TabIndex        =   154
                  Top             =   360
                  Width           =   615
               End
               Begin VB.TextBox TxtDay 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1560
                  TabIndex        =   153
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   315
                  Index           =   41
                  Left            =   840
                  TabIndex        =   159
                  Top             =   120
                  Width           =   405
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”‰…"
                  Height          =   315
                  Index           =   42
                  Left            =   120
                  TabIndex        =   158
                  Top             =   120
                  Width           =   405
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÌÊ„"
                  Height          =   315
                  Index           =   43
                  Left            =   1680
                  TabIndex        =   157
                  Top             =   120
                  Width           =   285
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·Œœ„…"
                  ForeColor       =   &H00C00000&
                  Height          =   555
                  Index           =   59
                  Left            =   2280
                  TabIndex        =   156
                  Top             =   0
                  Width           =   405
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   735
               Index           =   4
               Left            =   3840
               TabIndex        =   160
               TabStop         =   0   'False
               Top             =   2880
               Width           =   2985
               _cx             =   5265
               _cy             =   1296
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
               Begin VB.TextBox txtToOutSal 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0FF&
                  Height          =   285
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   164
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   675
               End
               Begin VB.TextBox TxtVSa 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0FF&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   163
                  Top             =   360
                  Width           =   675
               End
               Begin VB.TextBox TxtYaerOut 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0FF&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   162
                  Top             =   360
                  Width           =   675
               End
               Begin VB.TextBox TxtMontOut 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0FF&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   780
                  RightToLeft     =   -1  'True
                  TabIndex        =   161
                  Top             =   360
                  Width           =   675
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÌÊ„"
                  Height          =   315
                  Index           =   44
                  Left            =   1560
                  TabIndex        =   168
                  Top             =   120
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   315
                  Index           =   48
                  Left            =   990
                  TabIndex        =   167
                  Top             =   120
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”‰…"
                  Height          =   315
                  Index           =   49
                  Left            =   120
                  TabIndex        =   166
                  Top             =   120
                  Width           =   585
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã«“… »œÊ‰ —« »"
                  ForeColor       =   &H00C00000&
                  Height          =   675
                  Index           =   60
                  Left            =   2340
                  TabIndex        =   165
                  Top             =   0
                  Width           =   585
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   735
               Index           =   13
               Left            =   9960
               TabIndex        =   182
               TabStop         =   0   'False
               Top             =   3600
               Width           =   2505
               _cx             =   4419
               _cy             =   1296
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
               Begin VB.TextBox TxtYear2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Enabled         =   0   'False
                  Height          =   240
                  Left            =   90
                  TabIndex        =   185
                  Top             =   285
                  Width           =   435
               End
               Begin VB.TextBox TxtMonth2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Enabled         =   0   'False
                  Height          =   240
                  Left            =   615
                  TabIndex        =   184
                  Top             =   285
                  Width           =   465
               End
               Begin VB.TextBox TxtDay2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Enabled         =   0   'False
                  Height          =   240
                  Left            =   1125
                  TabIndex        =   183
                  Top             =   285
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   255
                  Index           =   64
                  Left            =   615
                  TabIndex        =   189
                  Top             =   90
                  Width           =   285
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”‰…"
                  Height          =   255
                  Index           =   65
                  Left            =   90
                  TabIndex        =   188
                  Top             =   90
                  Width           =   300
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÌÊ„"
                  Height          =   255
                  Index           =   66
                  Left            =   1230
                  TabIndex        =   187
                  Top             =   90
                  Width           =   195
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œ… «·⁄„·"
                  ForeColor       =   &H00C00000&
                  Height          =   450
                  Index           =   38
                  Left            =   1665
                  TabIndex        =   186
                  Top             =   0
                  Width           =   630
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   735
               Index           =   14
               Left            =   2880
               TabIndex        =   194
               TabStop         =   0   'False
               Top             =   3600
               Width           =   2505
               _cx             =   4419
               _cy             =   1296
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
               Begin VB.TextBox TxtDay3 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1185
                  TabIndex        =   197
                  Top             =   285
                  Width           =   480
               End
               Begin VB.TextBox TxtMonth3 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   645
                  TabIndex        =   196
                  Top             =   285
                  Width           =   480
               End
               Begin VB.TextBox TxtYear3 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   90
                  TabIndex        =   195
                  Top             =   285
                  Width           =   480
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "’«›Ì „œ… «·⁄„·"
                  ForeColor       =   &H00C00000&
                  Height          =   450
                  Index           =   71
                  Left            =   1740
                  TabIndex        =   201
                  Top             =   0
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÌÊ„"
                  Height          =   255
                  Index           =   72
                  Left            =   1290
                  TabIndex        =   200
                  Top             =   90
                  Width           =   210
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”‰…"
                  Height          =   255
                  Index           =   73
                  Left            =   90
                  TabIndex        =   199
                  Top             =   90
                  Width           =   315
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Â—"
                  Height          =   255
                  Index           =   74
                  Left            =   645
                  TabIndex        =   198
                  Top             =   90
                  Width           =   300
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—’Ìœ „—Õ·"
               ForeColor       =   &H00000000&
               Height          =   450
               Index           =   56
               Left            =   8955
               TabIndex        =   193
               Top             =   3720
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               ForeColor       =   &H00C00000&
               Height          =   450
               Index           =   40
               Left            =   7440
               TabIndex        =   192
               Top             =   3720
               Width           =   330
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ì«„ „Œ’Ê„Â"
               ForeColor       =   &H00000000&
               Height          =   450
               Index           =   70
               Left            =   6480
               TabIndex        =   191
               Top             =   3720
               Width           =   930
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì «·«Ì«„ "
               Height          =   405
               Index           =   53
               Left            =   1320
               TabIndex        =   180
               Top             =   4440
               Visible         =   0   'False
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ì«„ «·«Ã«“… «·„” Õﬁ…"
               Height          =   435
               Index           =   54
               Left            =   3720
               TabIndex        =   179
               Top             =   4440
               Width           =   1605
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "’«›Ì «·€Ì«»"
               Height          =   435
               Index           =   55
               Left            =   6525
               TabIndex        =   178
               Top             =   4440
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ì«„ «·«Ã«“… ﬁ»· «·Œ’„"
               Height          =   405
               Index           =   50
               Left            =   1200
               TabIndex        =   175
               Top             =   3720
               Width           =   1725
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ì«„ «·€Ì«»"
               Height          =   435
               Index           =   51
               Left            =   11280
               TabIndex        =   174
               Top             =   4440
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "»œÊ‰ —« »"
               Height          =   435
               Index           =   52
               Left            =   9000
               TabIndex        =   173
               Top             =   4440
               Width           =   1005
            End
            Begin VB.Label lblm 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„œÌ—"
               Height          =   372
               Left            =   8880
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   1080
               Width           =   852
            End
            Begin VB.Label lbres 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "”»» «·«Ã«“…"
               Height          =   495
               Left            =   11400
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   5040
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2580
               Index           =   62
               Left            =   2325
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1200
               Width           =   540
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6480
            Index           =   9
            Left            =   15
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   15
            Width           =   12480
            _cx             =   22013
            _cy             =   11430
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
               Height          =   4890
               Left            =   3270
               MaxLength       =   4
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   1365
               Width           =   840
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   3390
               Left            =   4245
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1725
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   3390
               Index           =   67
               Left            =   2145
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   1725
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   3240
               Index           =   68
               Left            =   4110
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   2175
               Width           =   45
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
               Height          =   3900
               Index           =   69
               Left            =   2955
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   1725
               Width           =   270
            End
         End
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   12525
      _cx             =   22093
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
      Caption         =   "ÿ·» √Ã«“…  "
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
         Left            =   1185
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
         ButtonImage     =   "formvocationl.frx":09B9
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
         ButtonImage     =   "formvocationl.frx":0D53
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
         ButtonImage     =   "formvocationl.frx":10ED
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
         ButtonImage     =   "formvocationl.frx":1487
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
         Left            =   4440
         Picture         =   "formvocationl.frx":1821
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
         TabIndex        =   34
         Top             =   0
         Width           =   2205
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   8
      Left            =   0
      TabIndex        =   128
      Top             =   0
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   612
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   110
      Top             =   3600
      Width           =   6012
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   7320
         TabIndex        =   111
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   237371393
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   6840
         TabIndex        =   112
         Top             =   1080
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo CbEmpReplaceMent 
         Bindings        =   "formvocationl.frx":5489
         Height          =   288
         Left            =   120
         TabIndex        =   121
         Top             =   240
         Width           =   4692
         _ExtentX        =   8281
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
         Caption         =   "«·„œÌ—"
         Height          =   288
         Index           =   31
         Left            =   4680
         TabIndex        =   120
         Top             =   1080
         Width           =   1008
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬂ«·Ì› «· √‘Ì—…"
         Height          =   288
         Index           =   29
         Left            =   5040
         TabIndex        =   119
         Top             =   720
         Width           =   1008
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—« » «·«”«”Ì"
         Height          =   285
         Index           =   28
         Left            =   8280
         TabIndex        =   118
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
         Height          =   285
         Index           =   26
         Left            =   8640
         TabIndex        =   117
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "0"
         Height          =   285
         Index           =   24
         Left            =   6840
         TabIndex        =   116
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›…"
         Height          =   288
         Left            =   5640
         TabIndex        =   115
         Top             =   600
         Width           =   648
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„— »…"
         Height          =   285
         Index           =   22
         Left            =   8520
         TabIndex        =   114
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ› «·»œÌ·"
         Height          =   288
         Index           =   21
         Left            =   4800
         TabIndex        =   113
         Top             =   240
         Width           =   1008
      End
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   315
      Left            =   120
      TabIndex        =   139
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   237436929
      CurrentDate     =   38784
   End
   Begin ImpulseButton.ISButton Accredit 
      Height          =   390
      Left            =   120
      TabIndex        =   169
      Top             =   9120
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   688
      ButtonPositionImage=   1
      Caption         =   "«—”«· ··«⁄ „«œ"
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorButton     =   -2147483635
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   315
      Left            =   0
      TabIndex        =   204
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   237436929
      CurrentDate     =   38784
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   288
      Index           =   2
      Left            =   8676
      TabIndex        =   91
      Top             =   2316
      Width           =   1008
   End
   Begin VB.Label XPTxtCurrent1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   2040
      TabIndex        =   87
      Top             =   8820
      Width           =   495
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
      TabIndex        =   38
      Top             =   4770
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   120
      Top             =   4680
      Width           =   6015
   End
   Begin VB.Label lblBr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      Height          =   255
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   780
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   12810
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   285
      Index           =   4
      Left            =   11430
      TabIndex        =   26
      Top             =   750
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   285
      Index           =   3
      Left            =   11430
      TabIndex        =   25
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   8310
      TabIndex        =   24
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   276
      Index           =   8
      Left            =   9120
      TabIndex        =   23
      Top             =   8760
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   2520
      TabIndex        =   22
      Top             =   8880
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   312
      Index           =   6
      Left            =   960
      TabIndex        =   21
      Top             =   8880
      Width           =   972
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   210
      TabIndex        =   20
      Top             =   8820
      Width           =   495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·Œ“‰…"
      Height          =   285
      Index           =   0
      Left            =   13350
      TabIndex        =   19
      Top             =   2130
      Width           =   1005
   End
End
Attribute VB_Name = "formvocatinl"
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
Dim rdio As String
Dim mSaveWithOutMsg As Boolean
Function CheckVacation() As Boolean
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "select * from TblVocationEntitlements where NoOrder =" & val(XPTxtID.text) & " "
Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
CheckVacation = True
Else
CheckVacation = False
End If
End Function
Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
 If val(XPTxtID.text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
      
      
 
 
    SendTopost Me.Name, "Tblvocation", "Id", val(DcbDetpartment.BoundText), val(dcBranch.BoundText), val(XPTxtID.text), XPTxtID
    
   rs.Resync
   
   
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

'    Cn.CommitTrans
'    BeginTrans = False
'FillApprovedTable
    Retrive (val(Me.XPTxtID.text))
End Sub

Private Sub Cmd_Click(index As Integer)
    ' On Error GoTo ErrTrap
    Select Case index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            Me.Rdb1(71).value = True
 
            Grid2.Clear flexClearScrollable, flexClearEverything
            Grid2.rows = 1
            Me.DCboUserName.BoundText = user_id
            ' TxtPaymentCounts.text = 1
            dcBranch.BoundText = Current_branch
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
            
            If CheckVacation = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·Õ—ﬂ… „— »ÿ… »„” Õﬁ«  «·«Ã«“…"
                Else
                    MsgBox "Can Not be Edited this Process linked to Vacation leave entitlements"
                End If
                Exit Sub
            End If
            If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ·.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If
       
            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
 
            Dim Msg As String

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
            Dim idd As Double
            If CheckSettingsVacType() = True Then
                Dim Period As Double
                Dim Diff   As Double
                ' Period = GetSettingsVacPeriod()
                ' DIFF = DateDiff("M", lastHolidaydate.value, xpdtbfrom.value)
                Diff = DateDiff("M", lastHolidaydate.value, xpdtbfrom.value)
                If CheckSettingsLikeContract() = True Then
                    GetHoldayDays2 val(DcboEmpName.BoundText), Period
                    Diff = val(TxtYear2.text) * 12 * 30 + val(TxtMonth2.text) * 30 + val(TxtDay2.text)
                    ' Diff = DateDiff("d", lastHolidaydate.value, xpdtbfrom.value)
    
                    Diff = Diff + GetLastBalanceMonthVaction(val(DcboEmpName.BoundText), val(XPTxtID.text)) * 30 - (val(TxtDiscouDay.text))
                    Period = Period * 30
                    ''//////////
                Else
                    Period = GetSettingsVacPeriod()
                End If
   
                If (Diff) >= Period Then
                Else
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "„œ… «·Œœ„… «ﬁ· „‰ «·„œ… «·„”„ÊÕ… ›Ì «⁄œ«œ«  «·«Ã«“« "
                    Else
                        MsgBox "The duration of service is less than the permitted period in the leave settings"
                    End If
                    Exit Sub
                End If
                DTPicker3.value = DateAdd("M", Period, lastHolidaydate.value)
                If GetSettingsVacDate(DTPicker3.value, idd) = True Then
                Else
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "«· «—ÌŒ €Ì— „ÿ«»ﬁ ·«⁄œ«œ«  «·«Ã«“« "
                    Else
                        MsgBox "Date does not match vacation settings"
                    End If
                    Exit Sub
                End If

                If GetSettingsVacDateAllow(xpdtbfrom.value, idd) = True Then
                Else
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "«· «—ÌŒ €Ì— „ÿ«»ﬁ ·«⁄œ«œ«  «·«Ã«“« "
                    Else
                        MsgBox "Date does not match vacation settings"
                    End If
                    Exit Sub
                End If
            End If
            'If GetHoldayDays(val(DcboEmpName.BoundText)) = 0 Then
            'If SystemOptions.UserInterface = ArabicInterface Then
            'MsgBox "Ì—ÃÏ «· «ﬂœ „‰ ⁄ﬁœ «·„ÊŸ›"
            'Else
            'MsgBox "Please make sure to contract the employee"
            'End If
            'Exit Sub
            'End If

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
            If CheckVacation = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… „— »ÿ… »„” Õﬁ«  «·«Ã«“…"
                Else
                    MsgBox "Can Not Delete this Process linked to Vacation leave entitlements"
                End If
                Exit Sub
            End If
            If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·«Ì„ﬂ‰ «·Õ–›.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not delete.This process associated with approvals"
                End If
                Exit Sub
            End If
            Del_Trans

        Case 5
            Load FrmEmpVacationSearch
            FrmEmpVacationSearch.index = 0
            FrmEmpVacationSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 8
            'CalCulateParts
            
        Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text)
        
            End If
            
        Case 11
            
            On Error Resume Next
            ShowAttachments XPTxtID.text, "2312202001"
        
    End Select

    Exit Sub
ErrTrap:
End Sub


Public Function GeBalancetHoldayDays(EmpID As Integer)
  Dim sql As String
  Dim Rs1 As New ADODB.Recordset
  Dim i As Integer
   Dim NODiffDate As Integer
   NODiffDate = 0
  sql = "SELECT    * from TblEmpHolidaysDetails WHERE     (Emp_id = " & EmpID & ")"
  Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
  If Rs1.RecordCount > 0 Then
  For i = 1 To Rs1.RecordCount
If Not IsNull(Rs1("DateExpectedM").value) Then
                If Not IsNull(Rs1("todate").value) Then
                 NODiffDate = NODiffDate + (val(DateDiff("d", Rs1("DateExpectedM").value, Rs1("todate").value)) * -1)
                 End If
                 End If
               Rs1.MoveNext
Next i
Else
End If
' NoDays = NODiffDate
       GeBalancetHoldayDays = NODiffDate
End Function



Public Function GetHoldayDays(Optional EmpID As Integer = 0, Optional ByRef NoDays As Integer, Optional ByRef NoDaysSala As Integer, Optional ByRef Tiket As Double, Optional ByRef netDay As Double)
  Dim sql As String
  Dim rs2 As New ADODB.Recordset
  Dim HoldaType As Integer
  Dim HoldaNo As Double
  Dim PriodType As Integer
  Dim PriodNo As Double
  Dim PriodType2 As Integer
  Dim PriodNo2 As Double
  Dim NODiffDate As Integer
  Dim tktval As Double
  Dim tkno As Integer
  Set rs2 = New ADODB.Recordset
  sql = "SELECT    * from dbo.Contract WHERE     (Emp_id = " & EmpID & ")"
  rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
  If rs2.RecordCount > 0 Then
 HoldaNo = IIf(IsNull(rs2("Holiday_period_no").value), 0, rs2("Holiday_period_no").value)
 HoldaType = IIf(IsNull(rs2("Holiday_period").value), -1, rs2("Holiday_period").value)
 PriodNo = IIf(IsNull(rs2("Due_period_no").value), 0, rs2("Due_period_no").value)
 PriodType = IIf(IsNull(rs2("due_period").value), -1, rs2("due_period").value)
 PriodNo2 = IIf(IsNull(rs2("salary_period_no").value), 0, rs2("salary_period_no").value)
 PriodType2 = IIf(IsNull(rs2("salary_period").value), -1, rs2("salary_period").value)
 tkno = IIf(IsNull(rs2("no_of_Child_ticket").value), 0, rs2("no_of_Child_ticket").value)
    tktval = IIf(IsNull(rs2("TicketValue").value), 0, rs2("TicketValue").value)
    Tiket = tkno * tktval
If PriodType2 = 1 Then
PriodNo2 = PriodNo2 * 30
End If
If HoldaType = 1 Then
HoldaNo = HoldaNo * 30
End If
If PriodType = 0 Then
PriodNo = PriodNo * 30
ElseIf PriodType = 1 Then
PriodNo = PriodNo * 366
End If
'NODiffDate = DateDiff("d", LastVocatinDate.value, DateSta.value)
If PriodNo > 0 Then
If val(NODiffDate / 366) <= 1 Then
NoDaysSala = (PriodNo2 / PriodNo) * NODiffDate
NoDays = (HoldaNo / PriodNo) * NODiffDate
ElseIf val((NODiffDate) / 366) < 2 Then
NoDays = (HoldaNo / PriodNo) * 366
NoDaysSala = (PriodNo2 / PriodNo) * 366
Else
NoDays = (HoldaNo / PriodNo) * 732
NoDaysSala = (PriodNo2 / PriodNo) * 732
End If
Else
End If
netDay = (HoldaNo / PriodNo)

GetHoldayDays = NoDaysSala
Else
GetHoldayDays = 0
  End If
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

 MySQL = " SELECT  TblEmployee_2.NumEkama, TblEmployee_2.BignDateWork,TblEmployee_2.IssueDateH,TblEmployee_2.lastHolidaydate,TblEmployee_2.lastHolidaydateH, dbo.TblVocation.ID, dbo.TblVocation.RecordDate, dbo.TblVocation.BranchID, dbo.TblVocation.ProjectID, dbo.TblVocation.FromDate, dbo.TblVocation.JobID,"
  MySQL = MySQL & "                  dbo.TblVocation.ToDate, dbo.TblVocation.Reson, dbo.TblVocation.Phone, dbo.TblVocation.Telephone, dbo.TblVocation.OtherAdress, dbo.TblVocation.VocationType,"
                   MySQL = MySQL & " dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblVocation.EmpID, dbo.TblVocation.ManagerID, dbo.EmpGroupDep.GroupName,"
                   MySQL = MySQL & " dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblVocation.Adress, TblEmployee_2.Emp_Code, TblEmployee_2.Emp_Name,"
                   MySQL = MySQL & " TblEmployee_2.Emp_Namee, TblEmployee_2.Emp_Name1, TblEmployee_2.Emp_Name2, TblEmployee_2.Emp_Name3, TblEmployee_2.Emp_Name4,"
                   MySQL = MySQL & " dbo.TblEmployee.Emp_Name AS Manager_EmpName, dbo.TblEmployee.Emp_Namee AS Manager_EmpNameE, TblEmployee_2.Fullcode, dbo.TblVocation.ResumeWorkH,"
                   MySQL = MySQL & "  dbo.TblVocation.ResumeWork, dbo.TblVocation.OutAndBack, dbo.TblVocation.OutOnly, dbo.TblVocation.ForFamily, dbo.TblVocation.WithoutSalary,"
                  MySQL = MySQL & "   dbo.TblVocation.WithSalary, dbo.TblVocation.ManagerApprove, dbo.TblVocation.VistCostOnCompany, dbo.TblVocation.VistCostOnEmployee, dbo.TblVocation.VisaCost,"
                   MySQL = MySQL & "  dbo.TblVocation.ForEmployee, replacement.Emp_Name AS ReplacementName, TblEmployee_2.BignDateWork, TblEmployee_2.SalaryInstrunse,"
                   MySQL = MySQL & "  TblEmployee_2.Nationality , dbo.TblVocation.TypeVacation , dbo.TblVocation.NoVacation ,dbo.TblVocation.notok , dbo.TblVocation.ok,"
                   
                    

MySQL = MySQL & " sign0 = " & GetUserSign(val(XPTxtID.text), Me.Name, 0)
MySQL = MySQL & " ,sign1 = " & GetUserSign(val(XPTxtID.text), Me.Name, 1)
MySQL = MySQL & " ,sign2 = " & GetUserSign(val(XPTxtID.text), Me.Name, 2)
MySQL = MySQL & " ,sign3 = " & GetUserSign(val(XPTxtID.text), Me.Name, 3)
MySQL = MySQL & " ,sign4 = " & GetUserSign(val(XPTxtID.text), Me.Name, 4)

 MySQL = MySQL & "  FROM     dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
                   MySQL = MySQL & "  dbo.TblVocation LEFT OUTER JOIN"
                   MySQL = MySQL & "  dbo.TblEmployee ON dbo.TblVocation.ManagerID = dbo.TblEmployee.Emp_ID ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblVocation.JobID LEFT OUTER JOIN"
                   MySQL = MySQL & "  dbo.EmpGroupDep ON dbo.TblVocation.ProjectID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
                   MySQL = MySQL & "  dbo.TblEmployee AS TblEmployee_2 ON dbo.TblVocation.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
                   MySQL = MySQL & "  dbo.TblBranchesData ON dbo.TblVocation.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
                   MySQL = MySQL & "  dbo.TblEmployee AS replacement ON dbo.TblVocation.EmpRemplacement = dbo.TblEmployee.Emp_ID"
 
       
       MySQL = MySQL & " Where (dbo.TblVocation.id = " & val(XPTxtID.text) & ")"
        
        
 Dim ii, jj
 ii = GeBalancetHoldayDays(val(DcboEmpName.BoundText))
 jj = GetHoldayDays(val(DcboEmpName.BoundText))
 
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\empmovee.rpt"
          '  StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_VacationRequest.rpt"
        Else
           StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\empmovee.rpt"
          ' StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_VacationRequest.rpt"
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
        xReport.ParameterFields(1).AddCurrentValue "" & ii & ""
         xReport.ParameterFields(2).AddCurrentValue "" & jj & ""
                 xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
       '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        
    Else
  xReport.ParameterFields(1).AddCurrentValue "" & ii & ""
         xReport.ParameterFields(2).AddCurrentValue "" & jj & ""
         
                 xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
       '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        
    End If

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
         ''//////
   Dim xLogo As CRAXDRT.OLEObject
   Dim SqlT As String
   Dim i As Integer
   Dim EmpIDD As Long
   Dim xWidth As Integer
   Dim Rs4 As ADODB.Recordset
   Set Rs4 = New ADODB.Recordset
  SqlT = " SELECT        TOP (100) PERCENT dbo.TblUsers.Empid"
  SqlT = SqlT + "    FROM            dbo.ApprovalData INNER JOIN"
  SqlT = SqlT + "                      dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
  SqlT = SqlT + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (NOT (ApprovDate IS NULL)) AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
  SqlT = SqlT & " ORDER BY levelorder"
  Rs4.Open SqlT, Cn, adOpenStatic, adLockOptimistic, adCmdText
  xWidth = 300
  For i = 1 To Rs4.RecordCount
  EmpIDD = IIf(IsNull(Rs4("Empid").value), 0, Rs4("Empid").value)
            If Dir(App.path & "\" & SystemOptions.ImagesPath & "\sign" & EmpIDD & ".JPG") <> "" Then
          
            
           Set xLogo = xReport.Areas(1).Sections(1).AddPictureObject(App.path & "\" & SystemOptions.ImagesPath & "\sign" & EmpIDD & ".JPG", xWidth, 1700)
            xLogo.Width = 800
            xLogo.Height = 800
            xLogo.backcolor = vbWhite
            xLogo.BorderColor = 255
            xLogo.CloseAtPageBreak = True
           xWidth = xWidth + 1000
          End If
        Rs4.MoveNext
    Next i
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub cmdApi_Click()

    Dim Req As New WinHttp.WinHttpRequest
    Req.Open "get", APIURL & "/api/empdata/getdata", async:=False
    Req.setRequestHeader "Content-Type", "application/hal+json"
    Req.setRequestHeader "Accept", "text/*, application/hal+json, application/json"
    Req.send
    Dim s As String
    Dim EmpID As Integer
    Dim rsDummy As New ADODB.Recordset
    Dim p As Object
    
    Set p = JSON.parse(Req.responseText)
    
    If Not (p Is Nothing) Then
        If JSON.GetParserErrors <> "" Then
            MsgBox JSON.GetParserErrors, vbInformation, "Parsing Error(s) occured"
        Else
            If p.count > 0 Then
                
                Dim i As Integer
                frmEmpVacList.FG.rows = 1
                For i = 1 To p.count
                    Dim empDic As Dictionary
                    Set empDic = p(i)
                    mSaveWithOutMsg = True
                    If Not empDic Is Nothing Then
                        frmEmpVacList.FG.AddItem ""
                        Dim row As Integer
                        row = frmEmpVacList.FG.rows - 1
                       
'                        frmEmpVacList.FG.TextMatrix(row, frmEmpVacList.FG.ColIndex("Id")) = empDic("id")
'                        frmEmpVacList.FG.TextMatrix(row, frmEmpVacList.FG.ColIndex("Code")) = empDic("employeeCode")
'                        frmEmpVacList.FG.TextMatrix(row, frmEmpVacList.FG.ColIndex("Name")) = empDic("employeeName")
'                        frmEmpVacList.FG.TextMatrix(row, frmEmpVacList.FG.ColIndex("from")) = Replace(empDic("startDate"), "T00:00:00", "")
'                        frmEmpVacList.FG.TextMatrix(row, frmEmpVacList.FG.ColIndex("to")) = Replace(empDic("endDate"), "T00:00:00", "")
'                        frmEmpVacList.FG.TextMatrix(row, frmEmpVacList.FG.ColIndex("notes")) = empDic("notes")
'                        frmEmpVacList.FG.TextMatrix(row, frmEmpVacList.FG.ColIndex("Sal")) = empDic("chkSallary")
                        s = "Select * from tblFromWeb where OrderNo = " & val(empDic("id")) & " and TransType = 0"
                        Set rsDummy = New ADODB.Recordset
                        rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
                        If Not rsDummy.EOF Then
                            GoTo NextRow
                        
                        End If
                        GetEmployeeIDFromCode empDic("employeeCode"), EmpID
                        
                        rsDummy.AddNew
'                        chkWithSalary.value = empDic("chkSallary") 'IIf(frmEmpVacList.salType, 1, 0)
'                        chkWithoutSalary.value = IIf(Not empDic("chkSallary"), 1, 0)
'                        TxtReson = empDic("notes")
'                        XPDtbFrom = Replace(empDic("startDate"), "T00:00:00", "")
'                        xpdtbto = Replace(empDic("endDate"), "T00:00:00", "")
                        
                        rsDummy!EmployeeCode = empDic("employeeCode")
                        rsDummy!chkSallary = empDic("chkSallary")
                        
                        rsDummy!TransType = 0
                        rsDummy!StartDate = Replace(empDic("startDate"), "T00:00:00", "")
                        rsDummy!EndDate = Replace(empDic("endDate"), "T00:00:00", "")
                        rsDummy!notes = empDic("notes")
                        rsDummy!orderNo = empDic("id")
                       ' rsDummy! = empDic("employeeCode")
                        rsDummy.update
                        
                        Cmd_Click 0
                        'GetEmployeeIDFromCode empDic("employeeCode"), EmpID
                        DcboEmpName.BoundText = EmpID
                        DcboEmpName_Click (0)
                        
                        chkWithSalary.value = IIf(empDic("chkSallary"), 1, 0)  'IIf(frmEmpVacList.salType, 1, 0)
                        chkWithoutSalary.value = IIf(Not empDic("chkSallary"), 1, 0)
                        txtreson = empDic("notes") & ""
                        xpdtbfrom = Replace(empDic("startDate"), "T00:00:00", "")
                        xpdtbto = Replace(empDic("endDate"), "T00:00:00", "")
                        SaveData
                        Accredit_Click
                        
NextRow:
'                        cm
                    End If
                Next
                mSaveWithOutMsg = False
                MsgBox " „ «·Õ›Ÿ"
                
'                frmEmpVacList.code = ""
'                frmEmpVacList.mIndex = 0
'                frmEmpVacList.show 1
'                If frmEmpVacList.code <> "" Then
'                    Dim EmpID As Integer
'                    EmpID = val(frmEmpVacList.code)
'                    GetEmployeeIDFromCode frmEmpVacList.code, EmpID
'                    'EmpID = val(frmEmpVacList.code)
'                    DcboEmpName.BoundText = EmpID
'                    DcboEmpName_Click (0)
'                    chkWithSalary.value = IIf(frmEmpVacList.salType, 1, 0)
'                    chkWithoutSalary.value = IIf(Not frmEmpVacList.salType, 1, 0)
'                    txtreson = frmEmpVacList.notes
'                    xpdtbfrom = frmEmpVacList.FromDate
'                    xpdtbto = frmEmpVacList.ToDate
'                End If
            Else
                MsgBox "No Data"
            End If
          
        End If
    Else
        MsgBox "An error occurred parsing json "
    End If
        
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub
Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub
Private Sub dtpResumeWorkh_LostFocus()
'If Me.TxtModFlg.Text <> "R" Then
'  VBA.Calendar = vbCalGreg
'            dtpResumeWork.value = ToGregorianDate(dtpResumeWorkh.value)
' End If
End Sub


Private Sub Fromdateh_LostFocus()
If Me.TxtModFlg.text <> "R" Then
  VBA.Calendar = vbCalGreg
            xpdtbfrom.value = ToGregorianDate(fromdateH.value)
            CalDate
 End If
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub

Private Sub Rdb1_Click(index As Integer)
Me.Txtother.Visible = False
End Sub


Private Sub Rdb2_Click(index As Integer)
Me.Txtother.Visible = False
End Sub




Private Sub RdbOther_Click(index As Integer)
Me.Txtother.Visible = True
End Sub

Private Sub ToDateH_LostFocus()
If Me.TxtModFlg.text <> "R" Then
  VBA.Calendar = vbCalGreg
            xpdtbto.value = ToGregorianDate(todateH.value)
           CalDate
 End If
End Sub


Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
 
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub

 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 14
        Set FrmEmployeeSearch.RetrunFrm = Me

        FrmEmployeeSearch.show
  
    End If

End Sub

Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
    
   If Me.TxtModFlg = "R" Then Exit Sub
   
   TxtNoVacation.text = DateDiff("d", xpdtbfrom.value, xpdtbto.value)
    Dim StrSQL As String
      
        
        Dim IssueDate As Date
        Dim DepID As Double
        Dim specid As Double
        Dim JobTypeID As Double
        Dim gradeID As Double
        Dim Account_code2 As String
           Dim Account_code  As String
        Dim Balance As String
        Dim mangerid As Integer
        Dim proj As Integer
        Dim endContractPerMonth As Double
        Dim lastHolidaydate2 As Date
        Dim lastHolidaydateH2 As String
        Dim BignDateWork As Date
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth, , mangerid, , proj, , , , , , , BignDateWork, , , , , , , , , , , , , , , , , , , , , , lastHolidaydate2, lastHolidaydateH2
        
        '  WriteCustomerBalPublic Account_code2, Balance
          '  WriteCustomerBalPublic Account_Code, Balance
         lastHolidaydate.value = lastHolidaydate2
         lastHolidaydateH.value = lastHolidaydateH2
         lastHolidaydate = GETlASTiSSUEDATENew((val(DcboEmpName.BoundText)), , 1)
         lastHolidaydateH.value = ToHijriDate(lastHolidaydate.value)
        DBIssueDate.value = IssueDate
        DcmbManagerID.BoundText = mangerid
         DcboJobsType.BoundText = JobTypeID
         DcboEmpDepartments.BoundText = proj
        DcbDetpartment.BoundText = DepID
        BignDate.value = BignDateWork
        BignDateH.value = ToHijriDate(BignDate)
                dateval
 ChekVacation11
  '   lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
        
    'End If
    Dim netDay As Double
    Dim NoVation As Double
    Dim IDees As String
 GetHoldayDays val(DcboEmpName.BoundText), , , , netDay
 'GetNoDayUnpadiVacation val(DcboEmpName.BoundText), IDees, NoVation
NoVation = GetNoDayUnpadiVacation2(val(DcboEmpName.BoundText), 0)
TxtDiscouDay.text = GetNoDayUnpadiVacation2(val(DcboEmpName.BoundText), 1)
 TxtVSa.text = NoVation
 TxtWithOutSala1 = TxtVSa.text
 TxtToalAbsent.text = Round(netDay * val(val(Me.TxtWithOutSala1.text) + val(TxtNewAbsent.text)), 0)
CalDate
End Sub

 

Private Sub TxtSearchCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
   FrmEmployee.show
     FrmEmployee.Retrive val(DcboEmpName.BoundText)
 End If
End Sub


Private Sub TxtWithOutSala1_Change()
DcboEmpName_Change
'TxtDuVocation.text = val(TxtContDay.text) - val(TxtToalAbsent.text) - val(TxtWithOutSala1.text)
End Sub
Private Sub TxtContDay_Change()
'TxtTotalDay.text = val(TxtContDay.text) + val(TxtLastDayVoc.text) - val(TxtToalAbsent.text)
TxtDuVocation.text = val(TxtContDay.text) - val(TxtToalAbsent.text)
End Sub
Private Sub ChekVacation11()
Dim PeriodMonth As Double
Dim Period As Double
Dim HoldyNo As Double
Dim NODiffDate As Double
Dim TempValu As Double
   ' If CheckSettingsVacType() = True Then
   ' GetHoldayDays2 val(DcboEmpName.BoundText), Period, HoldyNo
   ' PeriodMonth = DateDiff("M", lastHolidaydate.value, xpdtbfrom.value)
   ' If Period <> 0 Then
   ' TxtContDay.Text = Round((PeriodMonth / Period) * HoldyNo, 0)
   ' End If
   ' Else
       If CheckSettingsVacType() = True Then
    
    GetHoldayDays2 val(DcboEmpName.BoundText), Period, HoldyNo
   ' PeriodMonth = DateDiff("d", lastHolidaydate.value, XPDtbFrom.value)
PeriodMonth = val(TxtYear2.text) * 12 * 30 + val(TxtMonth2.text) * 30 + val(TxtDay2.text)
    
    If Period <> 0 Then
    If CheckSettingsLikeContract() = True Then
    NODiffDate = PeriodMonth + (GetLastBalanceMonthVaction(val(DcboEmpName.BoundText), val(XPTxtID.text)) * 30) - val(TxtDiscouDay.text)
    Period = Period * 30
   If NODiffDate >= Period Then
   TempValu = NODiffDate \ Period
    TxtContDay.text = Round((HoldyNo / Period) * NODiffDate, 2) 'HoldyNo * TempValu
   Else
    TxtContDay.text = 0
   End If
   ' TxtContDay.Text = Round((PeriodMonth / Period) * HoldyNo, 0)
    Else
    TxtContDay.text = Round((PeriodMonth / (Period * 30)) * HoldyNo, 2)
    End If
    End If
    Else
  Dim StrSQL As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
StrSQL = " SELECT     TOP 100 PERCENT EmpID, SUM([Value]) AS Tota"
StrSQL = StrSQL & " From dbo.tblVacationData"
StrSQL = StrSQL & " WHERE     (EmpID = " & val(Me.DcboEmpName.BoundText) & ") AND (ExpectedacationDate <=" & SQLDate(xpdtbfrom.value, True) & ") AND (Status1 IS NULL)"
StrSQL = StrSQL & " GROUP BY EmpID"
  Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
  If Rs3.RecordCount > 0 Then
 TxtContDay.text = IIf(IsNull(Rs3("Tota").value), 0, Rs3("Tota").value)
 End If
  End If
End Sub
Private Sub xpdtbfrom_Change()
        If Me.TxtModFlg.text <> "R" Then
                  fromdateH.value = ToHijriDate(xpdtbfrom.value)
              CalDate
        End If
        dateval
 ChekVacation11
DcboEmpName_Click (0)
End Sub
Sub CalDate()
lbl(37).Caption = ""
lbl(37).backcolor = &HE2E9E9
If Me.TxtModFlg.text <> "R" Then
 TxtNoVacation.text = DateDiff("d", xpdtbfrom.value, xpdtbto.value) + 1
 dtpResumeWork.value = DateAdd("d", 1, xpdtbto.value)
 If val(TxtNoVacation.text) < 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 lbl(37).Caption = " «—ÌŒ «·«Ã«“… Œÿ«¡"
 Else
 lbl(37).Caption = "Date is wrong"
 End If
 lbl(37).backcolor = &HFF&
 End If
 End If
End Sub
 
Private Sub xpdtbto_Change()
        If Me.TxtModFlg.text <> "R" Then
             
                  todateH.value = ToHijriDate(xpdtbto.value)
              CalDate
        End If
End Sub

Private Sub XPDtbTrans_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""

End Sub

Private Sub Dcbranch_Click(Area As Integer)
 
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim GrdBack As ClsBackGroundPic

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
    AddTip
    Set Dcombos = New ClsDataCombos
    
     Dcombos.GetEmployees Me.CbEmpReplaceMent
     Dcombos.GetUsers Me.DCboUserName
     Dcombos.GetEmployees Me.DcboEmpName
     Dcombos.GetBranches Me.dcBranch
       Dcombos.GetEmployees Me.DcmbManagerID
       Dcombos.GetEmpDepartments Me.DcbDetpartment


    ' Dcombos.GetEmpDepartments Me.DcmbFromDepart
 'Dcombos.GetEmpDepartments Me.DcboEmpDepartments
 Dcombos.GetEmpLocations Me.DcboEmpDepartments
    Dcombos.GetEmployees Me.DcmbManagerID
     Dcombos.GetEmpJobsTypes Me.DcboJobsType

    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
    
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblVocation    Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.text = "R"
        
 XPBtnMove_Click 2

    


    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
  
   Me.Refresh
   cmdApi_Click
   Retrive
   
    Exit Sub

ErrTrap:
End Sub

Private Sub ChangeLang()
    With Grid2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
    End With
    lbl(38).Caption = "Work Period"
    lbl(65).Caption = "Year"
    lbl(64).Caption = "Month"
    lbl(66).Caption = "Day"
lbl(52).Caption = "Unpaid Vac"
lbl(55).Caption = "Net Absnce"
 lbl(50).Caption = "Day Befor Discount"
lbl(39).Caption = "Last Vacation"
lbl(36).Caption = "Begin Work"
Label3.Caption = "Management"
lbl(53).Caption = "Total Day"
lbl(43).Caption = "Day"
lbl(44).Caption = "Day"
lbl(45).Caption = "Day"
lbl(41).Caption = "Month"
lbl(47).Caption = "Month"
lbl(48).Caption = "Month"
lbl(46).Caption = "Year"
lbl(42).Caption = "Year"
lbl(49).Caption = "Year"
lbl(60).Caption = "Without Salary"
lbl(51).Caption = "Net Absnce"
Frame1.Caption = "Visa"
lbl(58).Caption = "Days Absence"
lbl(54).Caption = "Days leave Entitlement"
chkManagerApprove.Caption = "Manager Approve"
chkWithSalary.Caption = "Paid Vacation"
chkWithoutSalary.Caption = "Not Paid Vacation"
lbl(21).Caption = "Replacement Employee"
lbl(9).Caption = "Visa Cost "
chkVistCostOnCompany.Caption = "Comapny Load "
chkVistCostOnEmployee.Caption = "Employee Load"
chkForFamily.Caption = "For Family"
chkForEmployee.Caption = "For Employee"
ChkOutAndBack.Caption = "Exit and return"
chkOutOnly.Caption = "Exit Only"
lbl(32).Caption = "Date OF Commencement"
lbl(10).Caption = "Manager"
lbl(18).Caption = "Visa"
lbl(9).Caption = "Visa Cost "
lbl(34).Caption = "No Vacation"
lbl(59).Caption = "Service"



    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
  '  Label1.Visible = False
XPTab301.Caption = "Data"
lbl(56).Caption = "Bal. Trans."
lbl(70).Caption = "Dis. Trans."

lbl(40).Caption = "Month"

lbl(71).Caption = "Net wowk"
lbl(72).Caption = "Day"
lbl(74).Caption = "Month"
lbl(73).Caption = "Year"


    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    Me.Caption = "Vacation Request"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
   lbl(3).Caption = "Employee"
 Me.lblBr.Caption = "Branch"
  Me.lblm.Caption = "Manager"
  Me.lblde.Caption = "DataOf Employee"
    lbl(15).Caption = "Location"
  lbl(35).Caption = "Start Vacation"
     lbl(33).Caption = "End Vacation"
     Me.Rdb1(71).Caption = "Official"
     Me.Rdb2(72).Caption = "Important"
     Me.Rdbother(73).Caption = "Other"
     Me.Rdb1(71).RightToLeft = False
     Me.Rdb2(72).RightToLeft = False
     Me.Rdbother(73).RightToLeft = False
   ' Fra(0).Caption = "payments Method"
 Me.lbltype.Caption = "Type Of Vacation"
    Me.gb.Caption = "Means Of Communication"
 Me.lbladres.Caption = "Adress"
   Me.lblmo.Caption = "Mobile"
   Me.lblj.Caption = "Job"
   ' ChkSaleryDis.Caption = "Auto Discount"
 Me.lbltel.Caption = "Telephone"
    Me.lblother.Caption = " Other."
    Me.lbres.Caption = "Reson Vacation"

  lbl(8).Caption = "By"
 lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
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

'Private Sub TxtAdvanceValue_LostFocus()
  '  Dim StrSQL As String
  '  Dim Mytot As String
  '  Dim MySal As String
  '  Exit Sub
 '   Dim Myrs As New ADODB.Recordset
    'StrSQL =
   ' Myrs.Open "SELECT * From TblEmployee  where EmpID=" & val(DcboEmpName.BoundText), Cn, adOpenStatic, adLockReadOnly

   ' If Not Myrs.EOF And Not IsNull(Myrs!Emp_Salary) Then
    '    MySal = Myrs!Emp_Salary
     '   Mytot = val(MySal) * 5
'
      '  If val(TxtAdvanceValue.text) >= Mytot Then
        '    MsgBox "⁄›Ê« «·”·›…  ⁄œ  «·Õœ  «·„”„ÊÕ »Â ÊÂÊ 5 «÷⁄«› ﬁÌ„Â «·—« »  " & Chr(13) & "   —« » «·„ÊŸ›    " & MySal, vbOKOnly, App.Title
        '    Exit Sub
   
  '      End If
  '
  '  End If
   
'End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = "ÿ·» «Ã«“…"
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
            '        Me.Caption = "ÿ·» «Ã«“…( ÃœÌœ )"
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
          '  TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = "ÿ·» «Ã«“…(  ⁄œÌ· )"
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
         '   TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
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
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent1.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If




  XPTxtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
  TxtDay2.text = IIf(IsNull(rs("TxtDay2").value), 0, rs("TxtDay2").value)
  TxtMonth2.text = IIf(IsNull(rs("TxtMonth2").value), 0, rs("TxtMonth2").value)
  TxtYear2.text = IIf(IsNull(rs("TxtYear2").value), 0, rs("TxtYear2").value)
  TxtDay3.text = IIf(IsNull(rs("TxtDay3").value), 0, rs("TxtDay3").value)
  TxtMonth3.text = IIf(IsNull(rs("TxtMonth3").value), 0, rs("TxtMonth3").value)
  TxtYear3.text = IIf(IsNull(rs("TxtYear3").value), 0, rs("TxtYear3").value)
  TxtAddDay.text = IIf(IsNull(rs("TxtAddDay").value), 0, rs("TxtAddDay").value)
  TxtDiscouDay.text = IIf(IsNull(rs("TxtDiscouDay").value), 0, rs("TxtDiscouDay").value)
'//////////////////////////////////
 CbEmpReplaceMent.BoundText = IIf(IsNull(rs("EmpRemplacement").value), "", rs("EmpRemplacement").value)

chkManagerApprove.value = IIf(IsNull(rs("ManagerApprove").value) Or rs("ManagerApprove").value = False, 0, 1)
chkVistCostOnEmployee.value = IIf(IsNull(rs("VistCostOnEmployee").value) Or rs("VistCostOnEmployee").value = False, 0, 1)
chkVistCostOnCompany.value = IIf(IsNull(rs("VistCostOnCompany").value) Or rs("VistCostOnCompany").value = False, 0, 1)
chkWithSalary.value = IIf(IsNull(rs("WithSalary").value) Or rs("WithSalary").value = False, 0, 1)
chkWithoutSalary.value = IIf(IsNull(rs("WithoutSalary").value) Or rs("WithoutSalary").value = False, 0, 1)
 chkForFamily.value = IIf(IsNull(rs("ForFamily").value) Or rs("ForFamily").value = False, 0, 1)
 chkForEmployee.value = IIf(IsNull(rs("ForEmployee").value) Or rs("ForEmployee").value = False, 0, 1)
chkOutOnly.value = IIf(IsNull(rs("OutOnly").value) Or rs("OutOnly").value = False, 0, 1)
 ChkOutAndBack.value = IIf(IsNull(rs("OutAndBack").value) Or rs("OutAndBack").value = False, 0, 1)
dtpResumeWork.value = IIf(IsNull(rs("ResumeWork").value), Date, rs("ResumeWork").value)
dtpResumeWorkh.value = IIf(IsNull(rs("ResumeWorkH").value), "", rs("ResumeWorkH").value)

txtVisaCost.text = IIf(IsNull(rs("VisaCost").value) Or rs("VisaCost").value = False, 0, 1)

Me.DcbDetpartment.BoundText = IIf(IsNull(rs("DeptID").value), "", rs("DeptID").value)
   
   
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    xpdtbfrom.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    fromdateH.value = IIf(IsNull(rs("fromdateH").value), "", rs("fromdateH").value)
   todateH.value = IIf(IsNull(rs("todateH").value), "", rs("todateH").value)
   lastHolidaydate.value = IIf(IsNull(rs("lastHolidaydate").value), Date, rs("lastHolidaydate").value)
   lastHolidaydateH.value = IIf(IsNull(rs("lastHolidaydateH").value), ToHijriDate(Date), rs("lastHolidaydateH").value)
    
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcboEmpDepartments.BoundText = IIf(IsNull(rs("ProjectID").value), "", rs("ProjectID").value)
'    DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value
    DcmbManagerID.BoundText = IIf(IsNull(rs("ManagerID").value), "", rs("ManagerID").value)
    Me.DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    
    DcboJobsType.BoundText = IIf(IsNull(rs("JobID").value), "", rs("JobID").value)
   xpdtbto.value = IIf(IsNull(rs("ToDate").value), Date, rs("ToDate").value)

  ' lbl(23).Caption = IIf(IsNull(rs("basicSalary").value), "", rs("basicSalary").value)
   xptxtphone.text = IIf(IsNull(rs("Phone").value), "", rs("Phone").value)
   xptxttelephone.text = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
  xptxtother.text = IIf(IsNull(rs("OtherAdress").value), "", rs("OtherAdress").value)
  txtreson.text = IIf(IsNull(rs("reson").value), "", rs("Reson").value)
  TxtAdress.text = IIf(IsNull(rs("Adress").value), "", rs("Adress").value)
  Txtother.text = IIf(IsNull(rs("VocationType").value), "", rs("VocationType").value)

  If Not IsNull(rs("TypeVacation").value) Then
  If (rs("TypeVacation").value) = 0 Then
  Me.Rdb1(71).value = True
  ElseIf (rs("TypeVacation").value) = 1 Then
  Me.Rdb2(72).value = True
  ElseIf (rs("TypeVacation").value) = 2 Then
  Me.Rdbother(73).value = True
  End If
  End If
  '''///23082017
  TxtDayAbs.text = IIf(IsNull(rs("DayAbs").value), "", rs("DayAbs").value)
  TxtMoAbs.text = IIf(IsNull(rs("MoAbs").value), "", rs("MoAbs").value)
  TxtYearAbs.text = IIf(IsNull(rs("YearAbs").value), "", rs("YearAbs").value)
  TxtDay.text = IIf(IsNull(rs("NoDay").value), "", rs("NoDay").value)
  TxtMonth.text = IIf(IsNull(rs("NoMonth").value), "", rs("NoMonth").value)
  Txtyear.text = IIf(IsNull(rs("NoYear").value), "", rs("NoYear").value)
 ' txtDayOut.Text = IIf(IsNull(rs("DayOut").value), "", rs("DayOut").value)
  TxtMontOut.text = IIf(IsNull(rs("MontOut").value), "", rs("MontOut").value)
  TxtYaerOut.text = IIf(IsNull(rs("YaerOut").value), "", rs("YaerOut").value)
  TxtContDay.text = IIf(IsNull(rs("ContDay").value), "", rs("ContDay").value)
  TxtNewAbsent.text = IIf(IsNull(rs("NewAbsent").value), "", rs("NewAbsent").value)
  TxtWithOutSala1.text = IIf(IsNull(rs("WithoutSala1").value), "", rs("WithoutSala1").value)
  TxtToalAbsent.text = IIf(IsNull(rs("ToalAbsent").value), "", rs("ToalAbsent").value)
  TxtDuVocation.text = IIf(IsNull(rs("DuVocation").value), "", rs("DuVocation").value)
  TxtTotalDay.text = IIf(IsNull(rs("TotalDay").value), "", rs("TotalDay").value)
  BignDate.value = IIf(IsNull(rs("BignDate").value), Date, rs("BignDate").value)
  BignDateH.value = IIf(IsNull(rs("BignDateH").value), ToHijriDate(Date), rs("BignDateH").value)

  'rdio = IIf(IsNull(rs("VocationType").value), "", rs("VocationType").value)
'If rdio = Me.Rdb1(71).Caption Then
'  Me.Rdb1(71).value = True
'  Else
'If rdio = Rdb2(72).Caption Then
'  Me.Rdb2(72).value = True
'
'  Else
'  Me.Txtother.Visible = False
'  Rdbother(73).value = True
' Me.Txtother = rdio
'End If
'End If

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
       If IsNull(rs("posted").value) Then
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               Accredit.Enabled = True
  Else
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " sent to Approval   "
                                               End If
                                               Accredit.Enabled = False
   End If
   TxtNoVacation.text = IIf(IsNull(rs("NoVacation").value), "", (rs("NoVacation").value))
   
  '  Set RsDetails = New ADODB.Recordset
  '  StrSQL = "Select * From  TblVocation Where ID=" & val(XPTxtID.text)
  '  RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    

  '  RsDetails.Close
  '  Set RsDetails = Nothing
    
    fillapprovData
    
    
    
    XPTxtCurrent1.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!! "
            Else
            Msg = "Please Select Employee"
        End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboEmpName.SetFocus
             Sendkeys "{F4}"
            Exit Sub
        End If

        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("Tblvocation", "ID", "", True))
         
            rs.AddNew
        'ElseIf Me.TxtModFlg.text = "E" Then
 

        End If

rs("EmpRemplacement").value = IIf(CbEmpReplaceMent.BoundText <> "", val(CbEmpReplaceMent.BoundText), 0)
rs("ManagerApprove").value = chkManagerApprove.value
rs("VistCostOnEmployee").value = chkVistCostOnEmployee.value
rs("VistCostOnCompany").value = chkVistCostOnCompany.value
rs("WithSalary").value = chkWithSalary.value
rs("WithoutSalary").value = chkWithoutSalary.value
rs("ForFamily").value = chkForFamily.value
rs("ForEmployee").value = chkForEmployee.value
rs("OutOnly").value = chkOutOnly.value
rs("OutAndBack").value = ChkOutAndBack.value
rs("ResumeWork").value = dtpResumeWork.value
rs("ResumeWorkH").value = dtpResumeWorkh.value
rs("lastHolidaydate").value = lastHolidaydate.value
rs("lastHolidaydateH").value = lastHolidaydateH.value
        rs("VisaCost").value = val(txtVisaCost.text)
        rs("NoVacation").value = val(TxtNoVacation.text)
        rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
        rs("ID").value = val(XPTxtID.text)
        rs("RecordDate").value = XPDtbTrans.value
        rs("EmpID").value = Me.DcboEmpName.BoundText
        rs("FromDate").value = Me.xpdtbfrom.value
        rs("ToDate").value = Me.xpdtbto.value
        rs("FromDateH").value = Me.fromdateH.value
        rs("ToDateH").value = Me.todateH.value
        rs("managerID").value = val(Me.DcmbManagerID.BoundText)
        rs("ProjectID").value = val(Me.DcboEmpDepartments.BoundText)
        rs("Adress").value = TxtAdress.text
        rs("TxtDay2").value = IIf(Me.TxtDay2.text = "", 0, val(TxtDay2.text))
        rs("TxtMonth2").value = IIf(Me.TxtMonth2.text = "", 0, val(TxtMonth2.text))
        rs("TxtYear2").value = IIf(Me.TxtYear2.text = "", 0, val(TxtYear2.text))
        rs("TxtDay3").value = IIf(Me.TxtDay3.text = "", 0, val(TxtDay3.text))
        rs("TxtMonth3").value = IIf(Me.TxtMonth3.text = "", 0, val(TxtMonth3.text))
        rs("TxtYear3").value = IIf(Me.TxtYear3.text = "", 0, val(TxtYear3.text))
        rs("TxtAddDay").value = IIf(Me.TxtAddDay.text = "", 0, val(TxtAddDay.text))
        rs("TxtDiscouDay").value = IIf(Me.TxtDiscouDay.text = "", 0, val(TxtDiscouDay.text))

        If Me.Rdb1(71).value = True Then
         rs("TypeVacation").value = 0
         ' rs("VocationType").value = Rdb1(71).Caption
          ElseIf Me.Rdb2(72).value = True Then
          
           rs("TypeVacation").value = 1
         ' rs("VocationType").value = Rdb2(72).Caption
          ElseIf Rdbother(73).value = True Then
            rs("TypeVacation").value = 2
          rs("VocationType").value = Txtother.text
          Else
          rs("TypeVacation").value = Null
          End If
              
        rs("DeptID").value = IIf(Me.DcbDetpartment.BoundText = "", Null, Me.DcbDetpartment.BoundText)
     
       'rs("gradeID").value = val(Me.DcboSpecifications.BoundText)
        rs("JobID").value = val(Me.DcboJobsType.BoundText)
        rs("Phone").value = Me.xptxtphone.text
        rs("Telephone").value = Me.xptxttelephone.text
        rs("OtherAdress").value = Me.xptxtother.text
        rs("reson").value = Me.txtreson.text
     ''////23082017
    rs("DayAbs").value = val(TxtDayAbs.text)
     rs("MoAbs").value = val(TxtMoAbs.text)
     rs("YearAbs").value = val(TxtYearAbs.text)
     rs("NoDay").value = val(TxtDay.text)
     rs("NoMonth").value = val(TxtMonth.text)
     rs("NoYear").value = val(Txtyear.text)
    ' rs("DayOut").value = val(txtDayOut.Text)
     rs("MontOut").value = val(TxtMontOut.text)
     rs("YaerOut").value = val(TxtYaerOut.text)
     rs("ContDay").value = val(TxtContDay.text)
     rs("NewAbsent").value = val(TxtNewAbsent.text)
     rs("WithoutSala1").value = val(TxtWithOutSala1.text)
     rs("ToalAbsent").value = val(TxtToalAbsent.text)
     rs("DuVocation").value = val(TxtDuVocation.text)
     rs("TotalDay").value = val(TxtTotalDay.text)
     rs("BignDate").value = BignDate.value
     rs("BignDateH").value = BignDateH.value
     rs("UserID").value = val(Me.DCboUserName.BoundText)
      rs.update
        Cn.CommitTrans
        BeginTrans = False
'        RsDetails.Close
        Set RsDetails = Nothing
        XPTxtCurrent1.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    If mSaveWithOutMsg Then Exit Sub
        Select Case Me.TxtModFlg.text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
              Else
              Msg = "This is Record Already Save " & CHR(1)
              Msg = Msg & "You want enter another record"
            End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
             MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
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
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
     Else
     Msg = "Can Not Save " & CHR(13)
     Msg = Msg & " The insert incorrect values "
       Msg = Msg & "Please make sure the data, and then try again"
     End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry error douring save"
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "ID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

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
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
   Else
        Msg = "Confirm Delete"
    End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
    Deletepost Me.Name, "Tblvocation", "Id", val(DcbDetpartment.BoundText), val(dcBranch.BoundText), val(XPTxtID.text), XPTxtID
    
                rs.delete
               
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent1.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This process is not available, as it has no records "
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
   Msg = "Sorry erorr douring delete"
   End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub



Function FillApprovedTable()
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.Name & "')"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
            Dim UserID As Integer
Dim EmpID As Integer
 
    
    If Rs1.RecordCount > 0 Then
            currentdate = Now
            
                    
                        GetApprovalDepartement val(DcbDetpartment.BoundText), UserID, EmpID
            
            If UserID <> 0 Then
           '***************************************
                                 RSApproval.AddNew
                        RSApproval("ScreenName").value = Me.Name
                        RSApproval("levelo").value = 1
                       RSApproval("EmpID").value = UserID
                        RSApproval("levelorder").value = 1
                         RSApproval("currorder").value = 1
                          RSApproval("Transaction_ID").value = val(XPTxtID.text)
                          RSApproval("NoteSerial").value = XPTxtID.text
                        RSApproval("Transaction_Date").value = Date
                        
                          RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
                       RSApproval("SendTime").value = currentdate
        
                 
                                RSApproval("Currcursor").value = 1
                                 RSApproval("FromUser").value = user_name
                     
                        
                        RSApproval.update
              End If
              
            
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.Name
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
                   RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 And UserID = 0 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                Rs1.MoveNext
            Next i

    End If
    
    

End Function



Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
   Grid2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    Grid2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label11.Caption = "Approved"
                                 End If
                            Label11.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label11.Caption = "Currently required Approve"
                            End If
                 Label11.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 Grid2.rows = 1
    End If
RsDetails.Close

End Function

Sub dateval()
If Me.TxtModFlg.text <> "R" Then
 
   Dim astrSplitItems() As String
    Dim Result As String
    
 
     Dim diff_year As Integer
    Result = ExactAge(BignDate.value, xpdtbfrom.value)
If Result <> "" Then
    astrSplitItems = Split(Result, "-")
    Txtyear.text = astrSplitItems(0)
    TxtMonth.text = astrSplitItems(1)
    TxtDay.text = astrSplitItems(2)
 End If
 Result = ExactAge(lastHolidaydate.value, xpdtbfrom.value)
 If Result <> "" Then
    astrSplitItems = Split(Result, "-")
    TxtYear2.text = astrSplitItems(0)
    TxtMonth2.text = astrSplitItems(1)
    TxtDay2.text = astrSplitItems(2)
 End If
   TxtAddDay.text = GetLastBalanceMonthVaction(val(DcboEmpName.BoundText))
  DTPicker4.value = DateAdd("d", val(TxtDiscouDay.text), lastHolidaydate)
  DTPicker3.value = DateAdd("d", val(TxtAddDay.text) * 30, xpdtbfrom)
        Result = ExactAge(DTPicker4.value, DTPicker3.value)
 If Result <> "" Then
    astrSplitItems = Split(Result, "-")
    TxtYear3.text = astrSplitItems(0)
    TxtMonth3.text = astrSplitItems(1)
    TxtDay3.text = astrSplitItems(2)
  End If
End If
End Sub
Sub GetHoldayDays2(Optional EmpID As Integer = 0, Optional ByRef PriodNo As Double, Optional ByRef HoldaNo As Double)
  Dim sql As String
  Dim rs As New ADODB.Recordset
  Dim PriodType As Integer
  Dim HoldaType As Integer
 ' Dim PriodNo As Double
  sql = "SELECT    * from dbo.Contract WHERE     (Emp_id = " & EmpID & ")"
  rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
  If rs.RecordCount > 0 Then
 PriodNo = IIf(IsNull(rs("Due_period_no").value), 0, rs("Due_period_no").value)
 PriodType = IIf(IsNull(rs("due_period").value), -1, rs("due_period").value)
  HoldaNo = IIf(IsNull(rs("Holiday_period_no").value), 0, rs("Holiday_period_no").value)
 HoldaType = IIf(IsNull(rs("Holiday_period").value), -1, rs("Holiday_period").value)
 If HoldaType = 1 Then
HoldaNo = HoldaNo * 30
End If
If PriodType = 2 Then
PriodNo = PriodNo / 30
ElseIf PriodType = 1 Then
PriodNo = PriodNo * 12
End If
End If
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

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "ÿ·» «Ã«“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «Ã«“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «Ã«“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «Ã«“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «Ã«“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «Ã«“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «Ã«“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «Ã«“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «Ã«“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «Ã«“…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "ÿ·» «Ã«“…", 1, 15204351, -2147483630
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

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub



