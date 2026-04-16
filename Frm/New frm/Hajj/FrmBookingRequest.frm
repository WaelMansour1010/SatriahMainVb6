VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmBookingRequest 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·» ÕÃ“"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12315
   Icon            =   "FrmBookingRequest.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   12315
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9675
      Left            =   0
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   0
      Width           =   12315
      _cx             =   21722
      _cy             =   17066
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
      Begin VB.TextBox TxtHotelMadinh 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   6720
         Width           =   2400
      End
      Begin VB.TextBox TxtHotelJaddah 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   6720
         Width           =   2400
      End
      Begin VB.TextBox TxtHotelMakh 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   6720
         Width           =   2400
      End
      Begin C1SizerLibCtl.C1Elastic pnlGrid 
         Height          =   2580
         Left            =   120
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   4005
         Width           =   12135
         _cx             =   21405
         _cy             =   4551
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
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   2085
            Left            =   120
            TabIndex        =   76
            Top             =   150
            Width           =   11985
            _cx             =   21140
            _cy             =   3678
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   14871017
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   16776960
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
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmBookingRequest.frx":038A
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
         Begin ImpulseButton.ISButton Cmd1 
            Height          =   270
            Index           =   0
            Left            =   11280
            TabIndex        =   96
            Top             =   2280
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
            ButtonImage     =   "FrmBookingRequest.frx":0532
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd1 
            Height          =   270
            Index           =   1
            Left            =   9480
            TabIndex        =   97
            Top             =   2280
            Width           =   1290
            _ExtentX        =   2275
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
            ButtonImage     =   "FrmBookingRequest.frx":0ACC
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   630
         Left            =   120
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   8040
         Width           =   11805
         _cx             =   20823
         _cy             =   1111
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
         Begin MSDataListLib.DataCombo DcbUser 
            Height          =   315
            Left            =   4050
            TabIndex        =   83
            Top             =   105
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   270
            Index           =   20
            Left            =   6720
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   105
            Width           =   675
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   300
            Left            =   1545
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   120
            Width           =   705
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   300
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   120
            Width           =   360
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   195
            Index           =   2
            Left            =   2340
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   120
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   300
            Index           =   4
            Left            =   795
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   120
            Width           =   615
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   735
         Left            =   0
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   0
         Width           =   12300
         _cx             =   21696
         _cy             =   1296
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   22.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "ÿ·» ÕÃ“     "
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
         CaptionStyle    =   1
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
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   37
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBookingRequest.frx":1066
            ColorButton     =   -2147483634
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
            Height          =   345
            Index           =   2
            Left            =   90
            TabIndex        =   38
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBookingRequest.frx":1400
            ColorButton     =   -2147483634
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
            Height          =   345
            Index           =   1
            Left            =   1680
            TabIndex        =   39
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBookingRequest.frx":179A
            ColorButton     =   -2147483634
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
            Height          =   345
            Index           =   3
            Left            =   615
            TabIndex        =   40
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBookingRequest.frx":1B34
            ColorButton     =   -2147483634
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   570
            Index           =   25
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   0
            Width           =   2655
         End
      End
      Begin C1SizerLibCtl.C1Elastic pnlHeader 
         Height          =   3105
         Left            =   120
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   840
         Width           =   12120
         _cx             =   21378
         _cy             =   5477
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
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   9420
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   90
            Width           =   1350
         End
         Begin VB.TextBox TxtOrder 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   960
            Width           =   1485
         End
         Begin VB.TextBox TxtCompnyOut 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   600
            Width           =   4365
         End
         Begin VB.TextBox TxtCompnyIn 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6165
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   600
            Width           =   4605
         End
         Begin VB.TextBox TxtCusNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   960
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9990
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   615
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox EmpMbile 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6165
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   1380
            Width           =   1485
         End
         Begin VB.TextBox EmpName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8565
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   1365
            Width           =   2205
         End
         Begin VB.ComboBox DcbModelID 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FrmBookingRequest.frx":1ECE
            Left            =   90
            List            =   "FrmBookingRequest.frx":1ED8
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   2160
            Width           =   4395
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3480
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   1710
            Width           =   1005
         End
         Begin VB.TextBox ID 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   9420
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   90
            Width           =   1350
         End
         Begin VB.TextBox TxtGroupName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6165
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   1005
            Width           =   4605
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   750
            Left            =   12270
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   2520
            Visible         =   0   'False
            Width           =   1080
            _cx             =   1905
            _cy             =   1323
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
            Caption         =   "«·„‘—›"
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
            Begin VB.OptionButton emp 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ÊŸ›"
               Height          =   390
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   195
               Value           =   -1  'True
               Width           =   720
            End
            Begin VB.OptionButton other 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√Œ—Ï"
               Height          =   390
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   195
               Width           =   480
            End
         End
         Begin VB.TextBox Model 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   90
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   4395
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.TextBox VehicleNo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9255
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   2160
            Width           =   1515
         End
         Begin VB.TextBox FlightNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6165
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   1710
            Width           =   1485
         End
         Begin MSComCtl2.DTPicker SDate 
            Height          =   330
            Left            =   3525
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   90
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   96206851
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo BranchID 
            Height          =   315
            Left            =   90
            TabIndex        =   3
            Top             =   90
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo GroupID 
            Height          =   315
            Left            =   11640
            TabIndex        =   61
            Top             =   -540
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo CompanyID 
            Height          =   315
            Left            =   12390
            TabIndex        =   63
            Top             =   3405
            Visible         =   0   'False
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AirPortID 
            Height          =   315
            Left            =   90
            TabIndex        =   10
            Top             =   1380
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AirLineID 
            Height          =   315
            Left            =   8565
            TabIndex        =   11
            Top             =   1710
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker ArriveDate 
            Height          =   285
            Left            =   9255
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2595
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   96206851
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo ProgrammID 
            Height          =   315
            Left            =   90
            TabIndex        =   17
            Top             =   2595
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo OutClientID 
            Height          =   315
            Left            =   90
            TabIndex        =   14
            Top             =   1710
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo VehicleType 
            Height          =   315
            Left            =   6165
            TabIndex        =   19
            Top             =   2160
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker ArriveTime 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "h:mm:ss AMPM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            Height          =   315
            Left            =   6165
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2595
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   95420418
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo InClientID 
            Height          =   315
            Left            =   6165
            TabIndex        =   95
            Top             =   615
            Visible         =   0   'False
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo SeasonsID 
            Height          =   315
            Left            =   6165
            TabIndex        =   1
            Top             =   120
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «„— «· ‘€Ì·"
            Height          =   210
            Index           =   27
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   960
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   450
            Index           =   26
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   1680
            Width           =   1320
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ ««·»—‰«„Ã ·œÌ ·⁄„Ì·"
            Height          =   450
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   975
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   330
            Index           =   23
            Left            =   8550
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   120
            Width           =   750
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Â« ›"
            Height          =   330
            Left            =   7110
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   1395
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘—›"
            Height          =   315
            Left            =   10710
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   1365
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ«›·« "
            Height          =   405
            Index           =   18
            Left            =   7710
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   2160
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÊœÌ·"
            Height          =   450
            Index           =   14
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   2160
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·Õ«›·« "
            Height          =   345
            Index           =   13
            Left            =   10950
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   2160
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "“„‰ «·Ê’Ê·"
            Height          =   285
            Index           =   12
            Left            =   7695
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   2625
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·»—‰«„Ã"
            Height          =   450
            Index           =   11
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   2625
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Ê’Ê· "
            Height          =   345
            Index           =   10
            Left            =   10350
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   2625
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·—Õ·…"
            Height          =   330
            Index           =   9
            Left            =   7110
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   1770
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ŒÿÊÿ «·ÃÊÌ…"
            Height          =   345
            Index           =   7
            Left            =   10350
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1770
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÿ«—"
            Height          =   450
            Index           =   5
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   1395
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·‘—ﬂ…"
            Height          =   345
            Index           =   3
            Left            =   11670
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   3330
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ã„Ê⁄…"
            Height          =   345
            Index           =   1
            Left            =   10950
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1020
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—ﬂ… „‰ «·Œ«—Ã"
            Height          =   210
            Index           =   0
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   615
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—ﬂ… «·”⁄ÊœÌ…"
            Height          =   345
            Index           =   6
            Left            =   10830
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   615
            Width           =   1200
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   330
            Index           =   24
            Left            =   1950
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   90
            Width           =   1200
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·ÌÊ„"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   5055
            TabIndex        =   57
            Top             =   90
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÕÃ“ —ﬁ„"
            Height          =   345
            Index           =   8
            Left            =   10950
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   90
            Width           =   1080
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   750
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   8835
         Width           =   12075
         _cx             =   21299
         _cy             =   1323
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   510
            Index           =   0
            Left            =   10635
            TabIndex        =   48
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   900
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
            ButtonImage     =   "FrmBookingRequest.frx":1EE8
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
            Height          =   510
            Index           =   1
            Left            =   9435
            TabIndex        =   49
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   900
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
            ButtonImage     =   "FrmBookingRequest.frx":874A
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
            Height          =   510
            Index           =   2
            Left            =   8070
            TabIndex        =   50
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   900
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
            ButtonImage     =   "FrmBookingRequest.frx":EFAC
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
            Height          =   510
            Index           =   3
            Left            =   6675
            TabIndex        =   51
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   900
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
            ButtonImage     =   "FrmBookingRequest.frx":1580E
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
            Height          =   510
            Index           =   4
            Left            =   5295
            TabIndex        =   52
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   900
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
            ButtonImage     =   "FrmBookingRequest.frx":1C070
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
            Height          =   510
            Index           =   6
            Left            =   1380
            TabIndex        =   53
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   900
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
            ButtonImage     =   "FrmBookingRequest.frx":228D2
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   510
            Left            =   105
            TabIndex        =   54
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   900
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
            ButtonImage     =   "FrmBookingRequest.frx":4C4F4
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
            Height          =   510
            Index           =   7
            Left            =   3960
            TabIndex        =   55
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   900
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
            ButtonImage     =   "FrmBookingRequest.frx":52D56
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
            Height          =   510
            Index           =   9
            Left            =   2610
            TabIndex        =   56
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   900
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
            ButtonImage     =   "FrmBookingRequest.frx":595B8
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   780
         Left            =   120
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   7200
         Width           =   12015
         _cx             =   21193
         _cy             =   1376
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
         Begin VB.TextBox TxtReservNo 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   4440
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   2115
         End
         Begin XtremeSuiteControls.CheckBox ApproveFlag 
            Height          =   270
            Left            =   11625
            TabIndex        =   82
            Top             =   300
            Visible         =   0   'False
            Width           =   270
            _Version        =   786432
            _ExtentX        =   476
            _ExtentY        =   476
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbUser2 
            Height          =   315
            Left            =   7560
            TabIndex        =   29
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker ApproveTime 
            Height          =   315
            Left            =   2280
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   95420418
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker ApproveDate 
            Height          =   315
            Left            =   120
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   95420419
            CurrentDate     =   37140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·—Õ·…"
            Height          =   330
            Index           =   22
            Left            =   6180
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«· «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1305
            TabIndex        =   86
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Êﬁ "
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3465
            TabIndex        =   85
            Top             =   240
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   300
            Index           =   21
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " √ﬂÌœ «·ÕÃ“"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   19
            Left            =   10665
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   0
            Width           =   1320
         End
      End
      Begin MSDataListLib.DataCombo MekkaHotelID 
         Height          =   315
         Left            =   8280
         TabIndex        =   25
         Top             =   7080
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo JeddahHotelID 
         Height          =   315
         Left            =   4320
         TabIndex        =   26
         Top             =   7080
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo MadinaHotelID 
         Height          =   315
         Left            =   240
         TabIndex        =   27
         Top             =   7080
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "›‰œﬁ «·„œÌ‰…"
         Height          =   375
         Index           =   17
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "›‰œﬁ Ãœ…"
         Height          =   375
         Index           =   16
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   6720
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "›‰œﬁ ›Ï „ﬂ…"
         Height          =   375
         Index           =   15
         Left            =   10860
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   6720
         Width           =   1080
      End
   End
End
Attribute VB_Name = "FrmBookingRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim TTP As clstooltip

Private Sub AirLineID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
FlightNo.SetFocus
End If
End Sub

Private Sub AirPortID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
AirLineID.SetFocus
End If
End Sub
Function GetOrderNo(Optional OrdeNo As Double) As Double
Dim Sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Sql = " SELECT   NoteSerial1 "
Sql = Sql & " From dbo.tblbookingrequest2"
Sql = Sql & " WHERE     (OrdeNo = " & OrdeNo & ")"
Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetOrderNo = IIf(IsNull(Rs3("NoteSerial1").value), 0, Rs3("NoteSerial1").value)
Else
GetOrderNo = 0
End If
End Function
Private Sub ApproveFlag_Click()
If SystemOptions.UserInterface = ArabicInterface Then
DcbUser2.BoundText = user_id
ApproveTime.value = Time
ApproveDate.value = Date
End If
End Sub

Private Sub BranchID_Change()
   If Me.TxtModFlg.Text <> "R" Then
   TxtNoteSerial1.Text = ""
      If ChekSanNumber(val(BranchID.BoundText), 70) = True Then
          TxtNoteSerial1.Text = ""
      End If
      TxtNoteSerial1.Text = ""
   End If
End Sub

Private Sub BranchID_Click(Area As Integer)
BranchID_Change
End Sub

Private Sub BranchID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Text1.SetFocus
End If
End Sub

Private Sub Cmd_Click(Index As Integer)
'    On Error GoTo ErrTrap
    Select Case Index
        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
             BranchID.BoundText = Current_branch
            ID.Text = CStr(new_id("tblbookingrequest", "ID", "", True))
           emp.value = True
           Grid.Clear flexClearScrollable, flexClearEverything
           Grid.Rows = Grid.FixedRows + 1
           DcbUser.BoundText = user_id
           ArriveTime.value = Time
           lbl(25).Caption = ""
           RelodCombo 0
           SeasonsID.BoundText = GetMosim(0)
        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If
If CHeckOrder() = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·Õ—ﬂ… „— »ÿ… »«„—  ‘€Ì·"
Else
MsgBox "Can Not Modify This Is Record Linked Orders Operation"
End If
Exit Sub
End If
            TxtModFlg.Text = "E"
            Grid.Rows = Grid.Rows + 1
        Case 2
        If val(BranchID.BoundText) = 0 Or BranchID.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·›—⁄"
        Else
        MsgBox "Please Select Branch"
        End If
        BranchID.SetFocus
        Exit Sub
        End If
        
        If val(SeasonsID.BoundText) = 0 Or SeasonsID.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·„Ê”„"
        Else
        MsgBox "Please Select The Season"
        End If
        SeasonsID.SetFocus
        Exit Sub
        End If
        If val(OutClientID.BoundText) = 0 Or OutClientID.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ⁄„Ì·"
        Else
        MsgBox "Please Select Customer"
        End If
        OutClientID.SetFocus
        Exit Sub
        End If
        
        If TxtCompnyIn.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «·‘—ﬂ… «·”⁄ÊœÌ…"
        Else
        MsgBox "Please Enter Company"
        End If
        TxtCompnyIn.SetFocus
        Exit Sub
        End If
        If TxtCompnyOut.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «·‘—ﬂ… «·Œ«—ÃÌ…"
        Else
        MsgBox "Please Enter Company"
        End If
        TxtCompnyOut.SetFocus
        Exit Sub
        End If
        If val(AirLineID.BoundText) = 0 Or AirLineID.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·ŒÿÊÿ «·ÃÊÌ…"
        Else
        MsgBox "Please Select Airlines"
        End If
        AirLineID.SetFocus
        Exit Sub
        End If

       If val(AirPortID.BoundText) = 0 Or AirPortID.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·„ÿ«—"
        Else
        MsgBox "Please Select The Airport"
        End If
        AirPortID.SetFocus
        Exit Sub
        End If
       If val(ProgrammID.BoundText) = 0 Or ProgrammID.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·»—‰«„Ã"
        Else
        MsgBox "Please Select The Type Of Program"
        End If
        ProgrammID.SetFocus
        Exit Sub
        End If
        If val(VehicleType.BoundText) = 0 Or VehicleType.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·Õ«›·…"
        Else
        MsgBox "Please Select The Type Of Vehicle"
        End If
        VehicleType.SetFocus
        Exit Sub
        End If
        If val(DcbModelID.Text) = 0 Or DcbModelID.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «·„ÊœÌ·"
        Else
        MsgBox "Please Select Model"
        End If
        DcbModelID.SetFocus
        Exit Sub
        End If
        If TxtGroupName.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «”„ «·„Ã„Ê⁄… "
        Else
        MsgBox "Please Enter  Group Name"
        End If
        TxtGroupName.SetFocus
        Exit Sub
        End If
        If EmpName.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «”„ «·„‘—› "
        Else
        MsgBox "Please Enter  Supervisor  Name"
        End If
        EmpName.SetFocus
        Exit Sub
        End If
    If EmpName.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «”„ «·„‘—› "
        Else
        MsgBox "Please Enter  Supervisor  Name"
        End If
        EmpName.SetFocus
        Exit Sub
        End If
        If EmpMbile.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· —ﬁ„ «·Â« › "
        Else
        MsgBox "Please Enter  Phone Number"
        End If
        EmpMbile.SetFocus
        Exit Sub
        End If
       If VehicleNo.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«·  «·Õ«›·«  "
        Else
        MsgBox "Please Enter  The Number Of Vehicles"
        End If
        VehicleNo.SetFocus
        Exit Sub
        End If
      If ChekGrid() = False Then
      If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "Ì—ÃÏ «œŒ«· „”«— Ê«Õœ ⁄·Ï «·«ﬁ·"
      Else
        MsgBox "Please Eneter One Path At Least"
      End If
      Exit Sub
      End If
   Dim TxtNoteSerial1str As String

    If TxtNoteSerial1.Text = "" Then
     TxtNoteSerial1str = Voucher_coding(val(Me.BranchID.BoundText), SDate.value, 70, 70, , , , , , val(SeasonsID.BoundText))
                If TxtNoteSerial1str = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…  Õ—ﬂ…  ÃœÌœ…  ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If TxtNoteSerial1str = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  «·Õ—ﬂ… ÃœÌœ     ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    End If
                End If
    End If
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If
If CHeckOrder() = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… „— »ÿ… »«„—  ‘€Ì·"
Else
MsgBox "Can Not Delete This Is Record Linked Orders Operation"
End If
Exit Sub
End If
            Del_Action

        Case 5

        Case 6
                Unload Me
         Case 7
                print_report2
         Case 9
         Unload FrmSearch_Hajj
         FrmSearch_Hajj.SendForm = "BookingRequest"
         FrmSearch_Hajj.show
         
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub Cmd1_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow
Case 1
Grid.Clear flexClearScrollable, flexClearEverything
      Grid.Rows = 2
End Select
End If
End Sub
Private Sub RemoveGridRow()

    With Me.Grid
MsgBox .Row
        If .Row <= 0 Then
                .Rows = 2
        Exit Sub
        Else
        .RemoveItem .Row
        End If
    End With
End Sub
Function ChekGrid() As Boolean
With Me.Grid
      If .Rows <= 1 Then
      ChekGrid = False
     ElseIf .Rows >= 2 Then
     If .TextMatrix(1, .ColIndex("PathName")) = "" Then
     ChekGrid = False
     Else
     ChekGrid = True
     End If
     End If
   End With
End Function

Private Sub CmdAttach_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.name, True) = False Then
                Exit Sub
            End If
ShowAttachments ID.Text, "20911201601"
End Sub

Private Sub DcbModelID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
TxtHotelMakh.SetFocus
End If
End Sub

Private Sub EmpMbile_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
TxtCusNo.SetFocus
End If
End Sub

Private Sub EmpName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
EmpMbile.SetFocus
End If
End Sub

Private Sub FlightNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
ProgrammID.SetFocus
End If
End Sub

Private Sub Form_Activate()
'    txtid.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
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

Private Sub Fill_Combos()
 Dim Dcombos As ClsDataCombos
  Dim str As String
  
   Set Dcombos = New ClsDataCombos
   Dcombos.GetBranches BranchID
   Dcombos.GetCompany InClientID, 0, 0
   Dcombos.GetCompany OutClientID, 2, 0
   If SystemOptions.UserInterface = ArabicInterface Then
   str = "select ID, Name from tblcompaniesgroup"
   Else
   str = "select ID, NameE from tblcompaniesgroup"
   End If
   fill_combo GroupID, str
   If SystemOptions.UserInterface = ArabicInterface Then
   str = "select Id , Name from TblTourismCompanies "
   Else
   str = "select Id , NameE from TblTourismCompanies "
   End If
   fill_combo CompanyID, str
  If SystemOptions.UserInterface = ArabicInterface Then
   str = "select Id , name  from tblairlines"
  Else
  str = "select Id , nameE  from tblairlines"
  End If
   fill_combo AirLineID, str
   If SystemOptions.UserInterface = ArabicInterface Then
    str = "select id , name from TblAirport "
    Else
    str = "select id , nameE from TblAirport "
   End If
   fill_combo AirPortID, str
   If SystemOptions.UserInterface = ArabicInterface Then
   str = "select id ,name from TblProgrammTypes "
   Else
   str = "select id ,nameE from TblProgrammTypes "
   End If
   fill_combo ProgrammID, str
  If SystemOptions.UserInterface = ArabicInterface Then
  str = "select id , name from tblhotels"
  Else
  str = "select id , nameE from tblhotels"
  End If
  fill_combo MekkaHotelID, str
  fill_combo JeddahHotelID, str
  fill_combo MadinaHotelID, str
  If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblCompaniesGroup  "
   Else
   str = " select id , nameE from TblCompaniesGroup  "
  End If
  str = str & " where Omra_Hajj=0"
   fill_combo SeasonsID, str
   Dcombos.GetTblCarsDataGroup VehicleType, 1, True
   Dcombos.GetUsers Me.DcbUser
   Dcombos.GetUsers Me.DcbUser2
   ' Dcombos.getCountriesGovernments Me.inCity
End Sub


Private Sub Form_Load()
 '   On Error GoTo ErrTrap
Dim I As Integer
 
        Fill_Combos
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  ÿ·» ÕÃ“ "
    LogTextE = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""
    DcbModelID.Clear
For I = 2015 To 2100
DcbModelID.AddItem I
Next I
    

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 '   Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset

    Dim StrSQL As String
    StrSQL = ""
    
     If SystemOptions.usertype <> UserAdminAll Then
      
StrSQL = "SELECT  *  From tblbookingrequest    "
  Else
 StrSQL = "SELECT  *  From tblbookingrequest"
    End If
  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2

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

    If Me.TxtModFlg.Text <> "R" Then

        Select Case Me.TxtModFlg.Text

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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
   
    lbl(7).Caption = " Name En"
    lbl(3).Caption = " Name Ar"
    lbl(8).Caption = "Process No"
    lbl(0).Caption = "Minister No."
    Label3.Caption = "School Manager"
 
    lbl(2).Caption = "Current Record"
    lbl(4).Caption = "NO. Recordes"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
   CmdAttach.Caption = "Attachment"

lbl(9).Caption = "Last Contract"



End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  ÿ·»  ÕÃ“   "
    LogTextE = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub

 

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim StrAccountCode As String
Dim Msg As String
'  Dim rs As New ADODB.Recordset
Dim StrSQL As String
Dim ClsAcc As New ClsAccounts
Dim LngRow As Long
Dim Sql As String
Dim count As Integer
Dim rate As Double
 
    With Grid

     Select Case .ColKey(Col)
 Case "PathName"
                        StrAccountCode = .ComboData
                        .TextMatrix(Row, .ColIndex("PathID")) = StrAccountCode
                        
                        
             'Case "FromCity"
             '           StrAccountCode = .ComboData
             '           .TextMatrix(Row, .ColIndex("FromcityId")) = StrAccountCode
             '           Grid.Rows = Grid.Rows + 1
             'Case "ToCity"
             '            StrAccountCode = .ComboData
             '           .TextMatrix(Row, .ColIndex("tocityId")) = StrAccountCode
     End Select
     If .Row = .Rows - 1 Then
     .Rows = .Rows + 1
     End If
End With




End Sub


 



Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

With Grid

If TxtModFlg.Text = "R" Then
          .ComboList = ""
          Cancel = True
End If


Select Case .ColKey(Col)
    Case "Date"
         .ComboList = ""
          Cancel = True
    Case "Time"
            .ComboList = ""
            Cancel = True
            Case "Remark"
        .ComboList = ""
    End Select
 End With

End Sub

Private Sub Grid_Click()
Select Case Grid.ColKey(Grid.Col)

Case "Date"
           Unload FrmRegesterDateProject
            FrmRegesterDateProject.SendForm = "BookingRequest"
          FrmRegesterDateProject.show vbModal
Case "Time"
             Unload FrmRegesterDateProject
             FrmRegesterDateProject.SendForm = "BookingRequest"
             FrmRegesterDateProject.show vbModal
End Select
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)


'Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    
     With Grid
     
     Select Case .ColKey(Col)
     
    Case "PathName"
        Set Rs_Temp = New ADODB.Recordset
        If SystemOptions.UserInterface = ArabicInterface Then
          StrSQL = " Select id,Name From TblShrines  "
          Else
          StrSQL = " Select id,NameE From TblShrines  "
          End If
          Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          If SystemOptions.UserInterface = ArabicInterface Then
          StrComboList = Grid.BuildComboList(Rs_Temp, "Name", "ID")
          Else
          StrComboList = Grid.BuildComboList(Rs_Temp, "NameE", "ID")
          End If
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList
          
 
          
     '  Case "FromCity"
     '     Set Rs_Temp = New ADODB.Recordset
     '     StrSQL = " Select GovernmentID,GovernmentName From TblCountriesGovernments  "
     '     Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
     '     StrComboList = Grid.BuildComboList(Rs_Temp, "GovernmentName", "GovernmentID")
     '      If StrComboList <> "" Then
     '            StrComboList = "|" & StrComboList
     '      End If
     '     .ComboList = StrComboList
     '
     '    Case "ToCity"
     '      Set Rs_Temp = New ADODB.Recordset
     '     StrSQL = " Select GovernmentID,GovernmentName From TblCountriesGovernments  "
     '     Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
     '     StrComboList = Grid.BuildComboList(Rs_Temp, "GovernmentName", "GovernmentID")
     '      If StrComboList <> "" Then
     '            StrComboList = "|" & StrComboList
     '      End If
     ''     .ComboList = StrComboList
      '
     End Select
   End With
End Sub

Private Sub GroupID_Change()

Dim str As String
Set Rs_Temp = New ADODB.Recordset
Set CompanyID.RowSource = Rs_Temp
If SystemOptions.UserInterface = ArabicInterface Then
    str = " select  ID , Name   from TblTourismCompanies where GroupID   = " & val(GroupID.BoundText)
Else
    str = " Select ID , NameE   TblTourismCompanies where GroupID  = " & val(GroupID.BoundText)
End If
fill_combo CompanyID, str
CompanyID.Refresh

End Sub





Private Sub ID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
SDate.SetFocus
End If
End Sub

Private Sub InClientID_Change()
InClientID_Click (0)
End Sub

Private Sub InClientID_Click(Area As Integer)
   Dim fullcode As String
    GetCustomersDetail val(InClientID.BoundText), , fullcode, 1
    Text1.Text = fullcode
End Sub

Private Sub InClientID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Text2.SetFocus
End If
End Sub

Private Sub Model_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.Model.Text, 1)
End Sub

Private Sub OutClientID_Change()
OutClientID_Click (0)
End Sub

Private Sub OutClientID_Click(Area As Integer)
   Dim fullcode As String
    GetCustomersDetail val(OutClientID.BoundText), , fullcode, 1
    Text2.Text = fullcode
End Sub

Private Sub OutClientID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
TxtGroupName.SetFocus
End If
End Sub

Private Sub ProgrammID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
VehicleNo.SetFocus
End If
End Sub

Private Sub SDate_Change()
   If Me.TxtModFlg.Text <> "R" Then
   TxtNoteSerial1.Text = ""
      If ChekSanNumber(val(BranchID.BoundText), 70) = True Then
          TxtNoteSerial1.Text = ""
      End If
      TxtNoteSerial1.Text = ""
   End If
End Sub

Private Sub SDate_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
SeasonsID.SetFocus
End If
End Sub

Private Sub SeasonsID_Change()
   If Me.TxtModFlg.Text <> "R" Then
   TxtNoteSerial1.Text = ""
      If ChekSanNumber(val(BranchID.BoundText), 70) = True Then
          TxtNoteSerial1.Text = ""
      End If
      TxtNoteSerial1.Text = ""
   End If
End Sub

Private Sub SeasonsID_Click(Area As Integer)
SeasonsID_Change
End Sub

Private Sub SeasonsID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
BranchID.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim CUSTID As Integer
   If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text1.Text, 2
        InClientID.BoundText = CUSTID
    End If
    InClientID.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim CUSTID As Integer
 If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text2.Text, 1
        OutClientID.BoundText = CUSTID
        OutClientID.SetFocus
    End If
End Sub

Private Sub TxtCusNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
AirPortID.SetFocus
End If
End Sub

Private Sub TxtGroupName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
EmpName.SetFocus
End If
End Sub

Private Sub TxtHotelJaddah_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.TxtHotelMadinh.SetFocus
End If
End Sub

Private Sub TxtHotelMakh_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.TxtHotelJaddah.SetFocus
End If
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ÿ·» ÕÃ“"
            Else
                Me.Caption = "School  Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(9).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            ID.locked = True
      

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            
            pnlHeader.Enabled = False
            
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  ÿ·» ÕÃ“ ( ÃœÌœ )"
            Else
                Me.Caption = "Booking Request Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   ÿ·» ÕÃ“ ( ÃœÌœ )"
            Else
                Me.Caption = "Booking Request Data(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            ID.locked = True
            pnlHeader.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   ÿ·» ÕÃ“ «·ÕÃ“ (  ⁄œÌ· )"
            Else
                Me.Caption = "Booking Request Data(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            ID.locked = True
           pnlHeader.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 MySQL = " SELECT     dbo.TblBookingRequest.ApproveFlag, dbo.TblBookingRequest.ApproveDate, dbo.TblBookingRequest.ApproveTime, dbo.TblBookingRequest.GroupName, "
 MySQL = MySQL & "                      dbo.TblBookingRequest.ModelID, dbo.TblBookingRequest.CreationDate, dbo.TblBookingRequest.ID, dbo.TblBookingRequest.SDate,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBookingRequest.InClientID,"
 MySQL = MySQL & "                      TblCustemers_2.CusName, TblCustemers_2.CusNamee, TblCustemers_2.Fullcode, dbo.TblBookingRequest.OutClientID, TblCustemers_1.CusName AS OutCusName,"
 MySQL = MySQL & "                      TblCustemers_1.CusNamee AS OutCusNameE, TblCustemers_1.Fullcode AS OutFullcode, dbo.TblBookingRequest.AirPortID, dbo.TblAirport.Name,"
 MySQL = MySQL & "                      dbo.TblAirport.NameE, dbo.TblBookingRequest.AirLineID, dbo.TblAirlines.Name AS AirLineName, dbo.TblAirlines.NameE AS AirLineNameE,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.EmpName, dbo.TblBookingRequest.EmpCode, dbo.TblBookingRequest.EmpMbile, dbo.TblBookingRequest.ArriveDate,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.ArriveTime, dbo.TblBookingRequest.emp, dbo.TblBookingRequest.other, dbo.TblBookingRequest.FlightNo, dbo.TblBookingRequest.UserID,"
 MySQL = MySQL & "                      TblUsers_2.UserName, dbo.TblBookingRequest.UserID2, TblUsers_1.UserName AS UserName2, dbo.TblBookingRequest.ProgrammID,"
 MySQL = MySQL & "                      dbo.TblProgrammTypes.Name AS ProgName, dbo.TblProgrammTypes.NameE AS ProgNameE, dbo.TblBookingRequest.VehicleNo,"
 MySQL = MySQL & "                      dbo.TBLCarTypes.name AS VehicTyname, dbo.TBLCarTypes.namee AS VehicTynameE, dbo.TblFlightDetails.[Date], dbo.TblFlightDetails.[Time],"
 MySQL = MySQL & "                      dbo.TblFlightDetails.Remarks, dbo.TblFlightDetails.PathID, dbo.TblShrines.Name AS PathName, dbo.TblShrines.NameE AS PathNameE,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.RemarkApprove, dbo.TblBookingRequest.StusID, dbo.TblBookingRequest.UseFlag, dbo.TblBookingRequest.ReservNo,"
 MySQL = MySQL & "                      dbo.TblCompaniesGroup.Name AS SeasoName, dbo.TblCompaniesGroup.NameE AS SeasoNameE, dbo.TblBookingRequest.SeasonsID,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.HotelMakh, dbo.TblBookingRequest.HotelMadinh, dbo.TblBookingRequest.HotelJaddah, dbo.TblBookingRequest.CusNo,"
 MySQL = MySQL & "                      dbo.TblBookingRequest.VehicleType , dbo.TblBookingRequest.CompnyIn, dbo.TblBookingRequest.CompnyOut , dbo.TblBookingRequest.NoteSerial1 "
 MySQL = MySQL & "  FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblBookingRequest LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblCompaniesGroup ON dbo.TblBookingRequest.SeasonsID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TBLCarTypes ON dbo.TblBookingRequest.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblProgrammTypes ON dbo.TblBookingRequest.ProgrammID = dbo.TblProgrammTypes.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblUsers TblUsers_1 ON dbo.TblBookingRequest.UserID2 = TblUsers_1.UserID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblUsers TblUsers_2 ON dbo.TblBookingRequest.UserID = TblUsers_2.UserID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblAirlines ON dbo.TblBookingRequest.AirLineID = dbo.TblAirlines.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblAirport ON dbo.TblBookingRequest.AirPortID = dbo.TblAirport.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblCustemers TblCustemers_1 ON dbo.TblBookingRequest.OutClientID = TblCustemers_1.CusID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblCustemers TblCustemers_2 ON dbo.TblBookingRequest.InClientID = TblCustemers_2.CusID ON"
 MySQL = MySQL & "                      dbo.TblBranchesData.branch_id = dbo.TblBookingRequest.BranchID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblShrines RIGHT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblFlightDetails ON dbo.TblShrines.ID = dbo.TblFlightDetails.PathID ON dbo.TblBookingRequest.ID = dbo.TblFlightDetails.HID"
 MySQL = MySQL & "  Where (dbo.TblBookingRequest.ID = " & val(ID.Text) & ")"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BookingRequest.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BookingRequest.rpt"
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
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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
Public Sub Retrive(Optional Lngid As Long = 0)

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
            rs.find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
    Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    TxtCompnyIn.Text = IIf(IsNull(rs("CompnyIn").value), "", Trim(rs("CompnyIn").value))
    TxtCompnyOut.Text = IIf(IsNull(rs("CompnyOut").value), "", Trim(rs("CompnyOut").value))
    TxtHotelMakh.Text = IIf(IsNull(rs("HotelMakh").value), "", (rs("HotelMakh").value))
    TxtHotelMadinh.Text = IIf(IsNull(rs("HotelMadinh").value), "", (rs("HotelMadinh").value))
    TxtHotelJaddah.Text = IIf(IsNull(rs("HotelJaddah").value), "", (rs("HotelJaddah").value))
    TxtCusNo.Text = IIf(IsNull(rs("CusNo").value), "", (rs("CusNo").value))
    ID.Text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    SDate.value = IIf(IsNull(rs("Sdate").value), Date, rs("Sdate").value)
    BranchID.BoundText = IIf(IsNull(rs("BranchID").value), "", Trim(rs("BranchID").value))
    InClientID.BoundText = IIf(IsNull(rs("InClientID").value), "", Trim(rs("InClientID").value))
    OutClientID.BoundText = IIf(IsNull(rs("OutClientID").value), "", Trim(rs("OutClientID").value))
    AirLineID.BoundText = IIf(IsNull(rs("AirLineID").value), "", Trim(rs("AirLineID").value))
    AirPortID.BoundText = IIf(IsNull(rs("AirPortID").value), "", Trim(rs("AirPortID").value))
    emp.value = IIf(IsNull(rs("emp").value), False, Trim(rs("emp").value))
    other.value = IIf(IsNull(rs("other").value), False, Trim(rs("other").value))
   ' EmpCode.text = IIf(IsNull(rs("EmpCode").value), "", Trim(rs("EmpCode").value))
    EmpName.Text = IIf(IsNull(rs("EmpName").value), "", Trim(rs("EmpName").value))
    EmpMbile.Text = IIf(IsNull(rs("EmpMbile").value), "", Trim(rs("EmpMbile").value))
    FlightNo.Text = IIf(IsNull(rs("FlightNo").value), "", Trim(rs("FlightNo").value))
    ArriveDate.value = IIf(IsNull(rs("ArriveDate").value), Date, Trim(rs("ArriveDate").value))
    ArriveTime.value = IIf(IsNull(rs("ArriveTime").value), Date, Trim(rs("ArriveTime").value))
    ProgrammID.BoundText = IIf(IsNull(rs("ProgrammID").value), "", Trim(rs("ProgrammID").value))
    VehicleNo.Text = IIf(IsNull(rs("VehicleNo").value), 0, Trim(rs("VehicleNo").value))
    Model.Text = IIf(IsNull(rs("Model").value), "", Trim(rs("Model").value))
    MekkaHotelID.BoundText = IIf(IsNull(rs("MekkaHotelID").value), "", Trim(rs("MekkaHotelID").value))
    MadinaHotelID.BoundText = IIf(IsNull(rs("MadinaHotelID").value), "", Trim(rs("MadinaHotelID").value))
    JeddahHotelID.BoundText = IIf(IsNull(rs("JeddahHotelID").value), "", Trim(rs("JeddahHotelID").value))
    VehicleType.BoundText = IIf(IsNull(rs("VehicleType").value), "", Trim(rs("VehicleType").value))
    GroupID.BoundText = IIf(IsNull(rs("GroupID").value), "", Trim(rs("GroupID").value))
    DcbModelID.Text = IIf(IsNull(rs("ModelID").value), 2016, Trim(rs("ModelID").value))
    TxtGroupName.Text = IIf(IsNull(rs("GroupName").value), "", Trim(rs("GroupName").value))
    Me.DcbUser.BoundText = IIf(IsNull(rs("UserID").value), "", Trim(rs("UserID").value))
    Me.DcbUser2.BoundText = IIf(IsNull(rs("UserID2").value), "", Trim(rs("UserID2").value))
    ApproveDate.value = IIf(IsNull(rs("ApproveDate").value), Date, Trim(rs("ApproveDate").value))
    TxtReservNo.Text = IIf(IsNull(rs("ReservNo").value), "", Trim(rs("ReservNo").value))
    SeasonsID.BoundText = IIf(IsNull(rs("SeasonsID").value), "", (rs("SeasonsID").value))
    TxtOrder.Text = GetOrderNo(val(ID.Text))
    If Not (IsNull(rs("ApproveFlag").value)) Then
    If rs("ApproveFlag").value = True Then
    ApproveFlag.value = vbChecked
    Else
    ApproveFlag.value = vbUnchecked
    End If
    End If
    If Not (IsNull(rs("StusID").value)) Then
    If rs("StusID").value = 1 Then
    ApproveFlag.value = vbChecked
    lbl(25).Caption = "ÕÃ“ „ƒﬂœ"
    ElseIf rs("StusID").value = 2 Then
    ApproveFlag.value = vbUnchecked
    lbl(25).Caption = "€Ì— „ƒﬂœ"
    Else
     lbl(25).Caption = "ÕÃ“ ÃœÌœ"
    ApproveFlag.value = vbUnchecked
    End If
    Else
    lbl(25).Caption = "ÕÃ“ ÃœÌœ"
    ApproveFlag.value = vbUnchecked
    End If
    Dim ContactTime As Date
     If Not IsNull(rs("ApproveTime").value) Then
     ContactTime = FormatDateTime(rs("ApproveTime").value, vbShortTime)
      Me.ApproveTime.value = ContactTime
    End If

    
    
    Set Rs_Temp = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT     dbo.TblFlightDetails.Remarks, dbo.TblFlightDetails.HID, dbo.TblFlightDetails.ID, dbo.TblFlightDetails.[Date], dbo.TblFlightDetails.[Time], dbo.TblFlightDetails.PathID, "
    StrSQL = StrSQL & "                  dbo.TblShrines.name , dbo.TblShrines.NameE"
    StrSQL = StrSQL & " FROM         dbo.TblFlightDetails LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblShrines ON dbo.TblFlightDetails.PathID = dbo.TblShrines.ID"
    StrSQL = StrSQL & "  where TblFlightDetails.HID = " & val(ID.Text)
    
    Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
     Rs_Temp.MoveFirst
     With Grid
        .Rows = Rs_Temp.RecordCount + 1
        Dim j As Integer
        For j = 1 To .Rows - 1
                .TextMatrix(j, .ColIndex("Serial")) = j
                .TextMatrix(j, .ColIndex("id")) = IIf(IsNull(Rs_Temp("id").value), "", Rs_Temp("id").value)
                .TextMatrix(j, .ColIndex("hid")) = IIf(IsNull(Rs_Temp("hid").value), 0, Rs_Temp("hid").value)
                '.TextMatrix(j, .ColIndex("fromcityid")) = IIf(IsNull(Rs_Temp("fromcity").value), "", Rs_Temp("fromcity").value)
                '.TextMatrix(j, .ColIndex("tocityid")) = IIf(IsNull(Rs_Temp("tocity").value), "", Rs_Temp("tocity").value)
                ' .TextMatrix(j, .ColIndex("fromcity")) = IIf(IsNull(Rs_Temp("fromcityname").value), "", Rs_Temp("fromcityname").value)
                '.TextMatrix(j, .ColIndex("tocity")) = IIf(IsNull(Rs_Temp("tocityname").value), "", Rs_Temp("tocityname").value)
                .TextMatrix(j, .ColIndex("Remark")) = IIf(IsNull(Rs_Temp("Remarks").value), "", Rs_Temp("Remarks").value)
                .TextMatrix(j, .ColIndex("date")) = IIf(IsNull(Rs_Temp("date").value), "", Rs_Temp("date").value)
                .TextMatrix(j, .ColIndex("time")) = IIf(IsNull(Rs_Temp("time").value), "", Rs_Temp("time").value)
                  .TextMatrix(j, .ColIndex("PathID")) = IIf(IsNull(Rs_Temp("PathID").value), 0, Rs_Temp("PathID").value)
                  If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(j, .ColIndex("PathName")) = IIf(IsNull(Rs_Temp("Name").value), "", Rs_Temp("Name").value)
                  Else
                    .TextMatrix(j, .ColIndex("PathName")) = IIf(IsNull(Rs_Temp("NameE").value), "", Rs_Temp("NameE").value)
                  End If
                Rs_Temp.MoveNext
         Next
        End With
    End If
    
    

    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub


 





Private Sub VehicleNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
VehicleType.SetFocus
End If
KeyAscii = KeyAscii_Num(KeyAscii, Me.VehicleNo.Text, 1)
End Sub

Private Sub VehicleType_Change()
If val(VehicleType.BoundText) <> 0 Then
RelodCombo val(VehicleType.BoundText)
End If
End Sub

Private Sub VehicleType_Click(Area As Integer)
VehicleType_Change
End Sub
Sub RelodCombo(Optional VehicleType As Integer)
   Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    If VehicleType = 0 Then
    Dcombos.ClearMyDataCombo ProgrammID
    Else
Dcombos.GetTblProgrammTypes ProgrammID, VehicleType
End If
End Sub
Private Sub VehicleType_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
DcbModelID.SetFocus
End If
End Sub
Function CHeckOrder() As Boolean
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim Sql As String
Sql = "select  * from tblbookingrequest where id=" & val(ID.Text) & " and UseFlag=1   "
Rs4.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
CHeckOrder = True
Else
CHeckOrder = False
End If
End Function
Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Grid.Rows = Grid.FixedRows
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
        Grid.Rows = Grid.FixedRows
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
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
    
        If Trim(BranchID.BoundText) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify Managerial Area"
            Else
                Msg = "Õœœ «·›—⁄ «Ê·« "
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            BranchID.SetFocus
   '         SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        Cn.BeginTrans
        BeginTrans = True

        Select Case Me.TxtModFlg.Text

           Case "N"
                rs.AddNew
                ID.Text = CStr(new_id("tblbookingrequest", "ID", "", True))
           Case "E"
                StrSQL = "delete From TblFlightDetails where  HID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 
           End Select

       
        
        
        rs("ID").value = val(ID.Text)
        If TxtNoteSerial1.Text = "" Then
              TxtNoteSerial1.Text = Voucher_coding(val(Me.BranchID.BoundText), SDate.value, 70, 70, , , , , , val(SeasonsID.BoundText))
        End If
        rs("NoteSerial1").value = IIf(Me.TxtNoteSerial1 <> "", val(TxtNoteSerial1.Text), Null)
          
        rs("SDate").value = SDate.value
        rs("BranchID").value = IIf(BranchID.BoundText = "", Null, BranchID.BoundText)
        rs("InClientID").value = IIf(InClientID.BoundText = "", Null, InClientID.BoundText)
        rs("OutClientID").value = IIf(OutClientID.BoundText = "", Null, OutClientID.BoundText)
        rs("GroupID").value = IIf(GroupID.BoundText = "", Null, GroupID.BoundText)
        rs("CompanyID").value = IIf(CompanyID.BoundText = "", Null, CompanyID.BoundText)
        rs("AirLineID").value = IIf(AirLineID.BoundText = "", Null, AirLineID.BoundText)
        rs("AirPortID").value = IIf(AirPortID.BoundText = "", Null, AirPortID.BoundText)
        rs("FlightNo").value = IIf(FlightNo.Text = "", Null, Trim(FlightNo.Text))
        rs("ArriveDate").value = ArriveDate.value
        rs("ArriveTime").value = FormatDateTime(ArriveTime.value, vbShortTime)
        rs("ReservNo").value = (TxtReservNo.Text)
        rs("ProgrammID").value = IIf(ProgrammID.BoundText = "", Null, (ProgrammID.BoundText))
        rs("VehicleNo").value = IIf(VehicleNo.Text = "", 0, val(VehicleNo.Text))
        rs("Model").value = IIf(Model.Text = "", 0, Model.Text)
       ' rs("MekkaHotelID").value = IIf(MekkaHotelID.BoundText = "", Null, (MekkaHotelID.BoundText))
       ' rs("JeddahHotelID").value = IIf(JeddahHotelID.BoundText = "", Null, (JeddahHotelID.BoundText))
       ' rs("MadinaHotelID").value = IIf(MadinaHotelID.BoundText = "", Null, (MadinaHotelID.BoundText))
       rs("CompnyIn").value = TxtCompnyIn.Text
       rs("CompnyOut").value = TxtCompnyOut.Text
       
        rs("HotelMakh").value = TxtHotelMakh.Text
        rs("HotelMadinh").value = TxtHotelMadinh.Text
        rs("HotelJaddah").value = TxtHotelJaddah.Text
        rs("CusNo").value = TxtCusNo.Text
        rs("VehicleType").value = IIf(VehicleType.BoundText = "", Null, (VehicleType.BoundText))
        'rs("EmpCode").value = IIf(EmpCode.text = "", Null, (EmpCode.text))
        rs("EmpName").value = IIf(EmpName.Text = "", Null, (EmpName.Text))
        rs("EmpMbile").value = IIf(EmpMbile.Text = "", Null, (EmpMbile.Text))
        rs("VehicleType").value = IIf(VehicleType.BoundText = "", Null, (VehicleType.BoundText))
        rs("emp").value = emp.value
        rs("other").value = other.value
        rs("FlightNo").value = FlightNo.Text
        rs("creationdate").value = Date
        rs("creationuserID").value = user_id
        rs("GroupID").value = IIf(GroupID.BoundText = "", Null, (GroupID.BoundText))
        rs("ModelID").value = IIf(val(DcbModelID.Text) = 0, Null, val(DcbModelID.Text))
        rs("CompanyID").value = IIf(CompanyID.BoundText = "", Null, (CompanyID.BoundText))
        rs("UserID").value = IIf(Me.DcbUser.BoundText = "", Null, val(DcbUser.BoundText))
       ' rs("UserID2").value = IIf(DcbUser2.BoundText = "", Null, val(DcbUser2.BoundText))
      '  rs("ApproveTime").value = FormatDateTime(ApproveTime.value, vbShortTime)
        rs("GroupName").value = IIf(Me.TxtGroupName.Text = "", Null, (TxtGroupName.Text))
       ' rs("ApproveDate").value = ApproveDate.value
        rs("SeasonsID").value = SeasonsID.BoundText
      '  If ApproveFlag.value = vbChecked Then
      '  rs("ApproveFlag").value = 1
      '  Else
      '  rs("ApproveFlag").value = 0
      '  End If
        rs.update
       Dim Rs_Temp As ADODB.Recordset
        Set Rs_Temp = New ADODB.Recordset
        StrSQL = " select * from TblFlightDetails  where 1 = -1 "
        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        With Grid
        Dim j As Integer
        For j = 1 To Grid.Rows - 1
           If val(.TextMatrix(j, .ColIndex("PathID"))) <> 0 Then
                    Rs_Temp.AddNew
                    Rs_Temp("ID") = CStr(new_id("TblFlightDetails", "ID", "", True))
                    Rs_Temp("HID") = val(ID.Text)
                    Rs_Temp("PathID") = val(.TextMatrix(j, .ColIndex("PathID")))
                   ' Rs_Temp("FromCity") = .TextMatrix(j, .ColIndex("FromCityid"))
                   ' Rs_Temp("ToCity") = .TextMatrix(j, .ColIndex("ToCityid"))
                    Rs_Temp("Date") = IIf(.TextMatrix(j, .ColIndex("Date")) = "", Null, .TextMatrix(j, .ColIndex("Date")))
                    Rs_Temp("Time") = IIf(.TextMatrix(j, .ColIndex("Time")) = "", Null, .TextMatrix(j, .ColIndex("Time")))
                    Rs_Temp("Remarks") = .TextMatrix(j, .ColIndex("Remark"))
                    Rs_Temp("creationdate").value = Date
                    Rs_Temp("creationuserID").value = user_id
                    Rs_Temp.update
                 End If
           Next
        End With
         
        
        
    
        Dim StrDes As String

     

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        'CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  ÿ·» ÕÃ“ " & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & Chr(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.Text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)
Grid.Rows = Grid.FixedRows
        Case "E"
            rs.find " ID=" & val(ID.Text) & "", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Action()
  
        Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If ID.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  ÿ·» ÕÃ“  —ﬁ„ " & Chr(13)
        Msg = Msg + (ID.Text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        Msg = "Delete Booking Request File ? " & Chr(13)
        Msg = Msg + (ID.Text) & Chr(13)
        Msg = Msg + "  Are you sure you want to delete ?"
        End If
        
        
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                                
                 StrSQL = "delete From TblFlightDetails where  HID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                           
                StrSQL = "delete From tblbookingrequest where  ID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 
                 rs.MoveFirst
                    
                   StrSQL = "SELECT  *  From tblbookingrequest "
                   rs.Close
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                   
                   
                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                    Grid.Rows = Grid.FixedRows
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Grid.Rows = Grid.FixedRows
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
         Msg = "this process Not Aailable"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… ÿ·» ÕÃ“ "
    Msg = Msg & Chr(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If

End Sub



Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«   √ﬂÌœ ÕÃ“ ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ÿ·» ÕÃ“" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ÿ·» ÕÃ“ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« ÿ·» ÕÃ“" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ÿ·» ÕÃ“" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  ÿ·» ÕÃ“", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub


