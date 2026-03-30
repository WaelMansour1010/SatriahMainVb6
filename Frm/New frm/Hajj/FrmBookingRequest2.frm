VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmBookingRequest2 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«„—  ‘€Ì·"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14085
   Icon            =   "FrmBookingRequest2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   14085
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9750
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   14085
      _cx             =   24844
      _cy             =   17198
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
      Begin VB.TextBox TxtHotelMakh 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9570
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   7470
         Width           =   3090
      End
      Begin VB.TextBox TxtHotelJaddah 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   7470
         Width           =   3105
      End
      Begin VB.TextBox TxtHotelMadinh 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   7470
         Width           =   3105
      End
      Begin C1SizerLibCtl.C1Elastic pnlGrid 
         Height          =   2760
         Left            =   0
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   4575
         Width           =   14025
         _cx             =   24739
         _cy             =   4868
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
            Height          =   2670
            Left            =   120
            TabIndex        =   38
            Top             =   0
            Width           =   13875
            _cx             =   24474
            _cy             =   4710
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
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmBookingRequest2.frx":038A
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
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   735
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   14070
         _cx             =   24818
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
         Caption         =   "«„—  ‘€Ì·"
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
            TabIndex        =   2
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   3
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
            ButtonImage     =   "FrmBookingRequest2.frx":0683
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
            TabIndex        =   4
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
            ButtonImage     =   "FrmBookingRequest2.frx":0A1D
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
            TabIndex        =   5
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
            ButtonImage     =   "FrmBookingRequest2.frx":0DB7
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
            TabIndex        =   6
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
            ButtonImage     =   "FrmBookingRequest2.frx":1151
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
      End
      Begin C1SizerLibCtl.C1Elastic pnlHeader 
         Height          =   3795
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Width           =   14010
         _cx             =   24712
         _cy             =   6694
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
         Begin VB.TextBox TxtVATNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   3480
            Width           =   3465
         End
         Begin VB.TextBox TxtFATYou 
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
            Height          =   300
            Left            =   6720
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Top             =   3120
            Width           =   1095
         End
         Begin VB.TextBox TxtFATValue 
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
            Height          =   300
            Left            =   4680
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   3120
            Width           =   975
         End
         Begin VB.TextBox TxtTotalValue 
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
            Height          =   300
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   133
            Top             =   3120
            Width           =   2775
         End
         Begin VB.TextBox TxtNoteSerialOrder 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   90
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox TxtDiscount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2085
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   2745
            Width           =   840
         End
         Begin VB.TextBox txtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   11670
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   120
            Width           =   870
         End
         Begin VB.TextBox TxtOreder2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   90
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtTotalNew 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9075
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   2745
            Width           =   1065
         End
         Begin VB.TextBox TxtCusNo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   90
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   2280
            Width           =   2835
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   7200
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   2355
            Width           =   645
         End
         Begin VB.TextBox TxtCompnyOut 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   90
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   585
            Width           =   2835
         End
         Begin VB.TextBox TxtCompnyIn 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9075
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   585
            Width           =   3465
         End
         Begin VB.TextBox TxtProAdd 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   6420
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   1740
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox TxtPathAddValue 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   11445
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   2745
            Width           =   1095
         End
         Begin VB.ComboBox DcbModelID 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FrmBookingRequest2.frx":14EB
            Left            =   9075
            List            =   "FrmBookingRequest2.frx":14F5
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   2355
            Width           =   3465
         End
         Begin VB.TextBox TxtTotal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9075
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   3105
            Width           =   3465
         End
         Begin VB.TextBox TxtNetDis 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   2745
            Width           =   840
         End
         Begin VB.ComboBox DcbTypeDis 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FrmBookingRequest2.frx":1505
            Left            =   4695
            List            =   "FrmBookingRequest2.frx":150F
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   2745
            Width           =   3150
         End
         Begin VB.TextBox TxtProgValue 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9075
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   1980
            Width           =   825
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   90
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   2355
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox TxtAirlounge 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   1605
            Width           =   2835
         End
         Begin VB.TextBox TXtMobile2 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   90
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   930
            Width           =   2835
         End
         Begin VB.TextBox TxtGroupName 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   9075
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   930
            Width           =   3465
         End
         Begin VB.TextBox EmpMbile 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4680
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   930
            Width           =   3165
         End
         Begin VB.TextBox EmpName 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   585
            Width           =   3165
         End
         Begin VB.TextBox TxtOreder 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   90
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   10365
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   585
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox ID 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   11670
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   120
            Width           =   870
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   555
            Left            =   14265
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   2040
            Visible         =   0   'False
            Width           =   3720
            _cx             =   6562
            _cy             =   979
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
               Height          =   510
               Left            =   15720
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   195
               Value           =   -1  'True
               Width           =   18570
            End
            Begin VB.OptionButton other 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√Œ—Ï"
               Height          =   510
               Left            =   3105
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   195
               Width           =   12645
            End
         End
         Begin VB.TextBox Model 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   -2055
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   3270
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.TextBox VehicleNo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   90
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   1980
            Width           =   2835
         End
         Begin VB.TextBox FlightNo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4680
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1275
            Width           =   3165
         End
         Begin MSComCtl2.DTPicker SDate 
            Height          =   300
            Left            =   9810
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99155971
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo BranchID 
            Height          =   315
            Left            =   5100
            TabIndex        =   11
            Top             =   120
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo InClientID 
            Height          =   315
            Left            =   9330
            TabIndex        =   13
            Top             =   585
            Visible         =   0   'False
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo GroupID 
            Height          =   315
            Left            =   9330
            TabIndex        =   16
            Top             =   930
            Visible         =   0   'False
            Width           =   3225
            _ExtentX        =   5689
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo CompanyID 
            Height          =   315
            Left            =   -1725
            TabIndex        =   18
            Top             =   1935
            Visible         =   0   'False
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AirPortID 
            Height          =   315
            Left            =   9075
            TabIndex        =   20
            Top             =   1275
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AirLineID 
            Height          =   315
            Left            =   90
            TabIndex        =   22
            Top             =   1275
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker ArriveDate 
            Height          =   300
            Left            =   9075
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1650
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "dd/mm/yyyy"
            Format          =   99155971
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo ProgrammID 
            Height          =   315
            Left            =   11205
            TabIndex        =   28
            Top             =   1995
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
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
            Height          =   270
            Left            =   4680
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1650
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   476
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99155970
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo VehicleType 
            Height          =   315
            Left            =   4680
            TabIndex        =   92
            Top             =   1980
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH 
            Height          =   300
            Left            =   8340
            TabIndex        =   95
            Top             =   105
            Width           =   1440
            _extentx        =   2725
            _extenty        =   582
         End
         Begin MSDataListLib.DataCombo OutClientID 
            Height          =   315
            Left            =   4680
            TabIndex        =   129
            Top             =   2355
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo SeasonsID 
            Height          =   315
            Left            =   2520
            TabIndex        =   130
            Top             =   120
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AccountVat 
            Bindings        =   "FrmBookingRequest2.frx":151F
            Height          =   315
            Left            =   240
            TabIndex        =   139
            Top             =   3360
            Visible         =   0   'False
            Width           =   3450
            _ExtentX        =   6085
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
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·÷—Ì» ··⁄„Ì·"
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   12480
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»…«·›« "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   66
            Left            =   8205
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   3120
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·›« "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   67
            Left            =   5685
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   3120
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ì"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   68
            Left            =   3525
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Top             =   3120
            Width           =   570
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   330
            Index           =   30
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   120
            Width           =   750
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·»—‰«„Ã"
            Height          =   300
            Left            =   10155
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   2745
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   345
            Index           =   27
            Left            =   7635
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   2355
            Width           =   1260
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·«÷«›Ì"
            Height          =   300
            Left            =   12570
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   2745
            Width           =   1245
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   13155
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   3105
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·Œ’„"
            Height          =   300
            Index           =   26
            Left            =   1065
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   2745
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·Œ’„"
            Height          =   300
            Index           =   25
            Left            =   3255
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   2745
            Width           =   900
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   300
            Left            =   7995
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   2745
            Width           =   900
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·»—‰«„Ã"
            Height          =   300
            Left            =   10155
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   1980
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·»—‰«„Ã ·œÌ «·⁄„Ì·"
            Height          =   345
            Index           =   23
            Left            =   2895
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   2235
            Width           =   1680
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«·…"
            Height          =   285
            Left            =   3300
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1605
            Width           =   1275
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÃÊ«·"
            Height          =   300
            Left            =   3300
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   930
            Width           =   1275
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Â« ›"
            Height          =   300
            Left            =   7635
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   930
            Width           =   1260
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘—›"
            Height          =   300
            Index           =   0
            Left            =   7635
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   585
            Width           =   1260
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «ﬂÌœ ÕÃ“ —ﬁ„"
            Height          =   195
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   120
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ«›·« "
            Height          =   300
            Index           =   18
            Left            =   7635
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   1980
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÊœÌ·"
            Height          =   345
            Index           =   14
            Left            =   12555
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   2355
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·Õ«›·« "
            Height          =   300
            Index           =   13
            Left            =   3300
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1995
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "“„‰ «·Ê’Ê·"
            Height          =   300
            Index           =   12
            Left            =   7635
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   1650
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»—‰«„Ã"
            Height          =   300
            Index           =   11
            Left            =   12555
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   1980
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Ê’Ê· "
            Height          =   300
            Index           =   10
            Left            =   12555
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   1650
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·—Õ·…"
            Height          =   300
            Index           =   9
            Left            =   7635
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1275
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ŒÿÊÿ «·ÃÊÌ…"
            Height          =   300
            Index           =   7
            Left            =   3300
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1275
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÿ«—"
            Height          =   300
            Index           =   5
            Left            =   12555
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1275
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·‘—ﬂ…"
            Height          =   285
            Index           =   3
            Left            =   -450
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   2055
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ã„Ê⁄…"
            Height          =   300
            Index           =   1
            Left            =   12555
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   930
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—ﬂ… »«·Œ«—Ã"
            Height          =   300
            Index           =   0
            Left            =   3285
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   600
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—ﬂ… «·”⁄ÊœÌ…"
            Height          =   300
            Index           =   6
            Left            =   12555
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   585
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   300
            Index           =   24
            Left            =   7185
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   120
            Width           =   945
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ "
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   10785
            TabIndex        =   10
            Top             =   120
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   315
            Index           =   8
            Left            =   12555
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   120
            Width           =   1260
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   765
         Left            =   120
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   10830
         Visible         =   0   'False
         Width           =   12105
         _cx             =   21352
         _cy             =   1349
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
         Begin XtremeSuiteControls.CheckBox ApproveFlag 
            Height          =   270
            Left            =   11145
            TabIndex        =   44
            Top             =   300
            Width           =   750
            _Version        =   786432
            _ExtentX        =   1323
            _ExtentY        =   476
            _StockProps     =   79
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbUser2 
            Height          =   315
            Left            =   6720
            TabIndex        =   45
            Top             =   240
            Width           =   3375
            _ExtentX        =   5953
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
            Left            =   3600
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   240
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99155970
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker ApproveDate 
            Height          =   315
            Left            =   360
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   240
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   99155971
            CurrentDate     =   37140
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«· «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2265
            TabIndex        =   50
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Êﬁ "
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5505
            TabIndex        =   48
            Top             =   240
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   300
            Index           =   21
            Left            =   9840
            RightToLeft     =   -1  'True
            TabIndex        =   46
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
            TabIndex        =   43
            Top             =   0
            Width           =   1320
         End
      End
      Begin MSDataListLib.DataCombo MekkaHotelID 
         Height          =   315
         Left            =   8130
         TabIndex        =   67
         Top             =   11535
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo JeddahHotelID 
         Height          =   315
         Left            =   4200
         TabIndex        =   69
         Top             =   11535
         Visible         =   0   'False
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo MadinaHotelID 
         Height          =   315
         Left            =   120
         TabIndex        =   71
         Top             =   11535
         Visible         =   0   'False
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Style           =   2
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   585
         Left            =   510
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   4785
         Width           =   12135
         _cx             =   21405
         _cy             =   1032
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
            Left            =   6900
            TabIndex        =   97
            Top             =   105
            Width           =   4050
            _ExtentX        =   7144
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
            Height          =   300
            Index           =   20
            Left            =   10725
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   105
            Width           =   1155
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   330
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   330
            Left            =   465
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   2
            Left            =   3945
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   330
            Index           =   4
            Left            =   1365
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   120
            Width           =   1035
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   555
         Index           =   11
         Left            =   120
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   7830
         Width           =   13935
         _cx             =   24580
         _cy             =   979
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
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   375
            Left            =   7905
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   240
            Width           =   3060
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9600
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSDataListLib.DataCombo DcbUserVoucher 
            Height          =   315
            Left            =   240
            TabIndex        =   108
            Top             =   120
            Width           =   4455
            _ExtentX        =   7858
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
            Caption         =   " „ «‰‘«¡ «·ﬁÌœ »Ê«”ÿ…"
            Height          =   255
            Index           =   28
            Left            =   4695
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   120
            Width           =   2550
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   255
            Index           =   35
            Left            =   12390
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   120
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   735
         Left            =   120
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   8910
         Width           =   13845
         _cx             =   24421
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   12570
            TabIndex        =   113
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
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
            ButtonImage     =   "FrmBookingRequest2.frx":1534
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
            Height          =   495
            Index           =   1
            Left            =   11145
            TabIndex        =   114
            Top             =   120
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   873
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
            ButtonImage     =   "FrmBookingRequest2.frx":7D96
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
            Height          =   495
            Index           =   2
            Left            =   9315
            TabIndex        =   115
            Top             =   120
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   873
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
            ButtonImage     =   "FrmBookingRequest2.frx":E5F8
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
            Height          =   495
            Index           =   3
            Left            =   8280
            TabIndex        =   116
            Top             =   120
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   873
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
            ButtonImage     =   "FrmBookingRequest2.frx":14E5A
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
            Height          =   495
            Index           =   4
            Left            =   7035
            TabIndex        =   117
            Top             =   120
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
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
            ButtonImage     =   "FrmBookingRequest2.frx":1B6BC
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
            Height          =   495
            Index           =   6
            Left            =   1365
            TabIndex        =   118
            Top             =   120
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   873
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
            ButtonImage     =   "FrmBookingRequest2.frx":21F1E
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
            Height          =   495
            Left            =   105
            TabIndex        =   119
            Top             =   120
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   873
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
            ButtonImage     =   "FrmBookingRequest2.frx":4BB40
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
            Height          =   495
            Index           =   7
            Left            =   5805
            TabIndex        =   120
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   873
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
            ButtonImage     =   "FrmBookingRequest2.frx":523A2
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
            Height          =   495
            Index           =   9
            Left            =   4965
            TabIndex        =   121
            Top             =   120
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   873
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
            ButtonImage     =   "FrmBookingRequest2.frx":58C04
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton BtnCreateVou 
            Height          =   495
            Left            =   3450
            TabIndex        =   122
            Top             =   120
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "«‰‘«¡ «·ﬁÌœ"
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
            ButtonImage     =   "FrmBookingRequest2.frx":5F466
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton BtnDeleteVoun 
            Height          =   495
            Left            =   2145
            TabIndex        =   123
            Top             =   120
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬁÌœ"
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
            ButtonImage     =   "FrmBookingRequest2.frx":65CC8
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
      Begin MSDataListLib.DataCombo DcbUser3 
         Height          =   315
         Left            =   6030
         TabIndex        =   124
         Top             =   8520
         Width           =   6780
         _ExtentX        =   11959
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
         Height          =   255
         Index           =   29
         Left            =   11115
         RightToLeft     =   -1  'True
         TabIndex        =   125
         Top             =   8520
         Width           =   2580
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "›‰œﬁ «·„œÌ‰…"
         Height          =   315
         Index           =   17
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   7470
         Width           =   1665
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "›‰œﬁ Ãœ…"
         Height          =   315
         Index           =   16
         Left            =   6780
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   7470
         Width           =   2265
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "›‰œﬁ ›Ï „ﬂ…"
         Height          =   315
         Index           =   15
         Left            =   12555
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   7470
         Width           =   1260
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·Õ«›·« "
      Height          =   315
      Index           =   22
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   0
      Width           =   1245
   End
End
Attribute VB_Name = "FrmBookingRequest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim TTP As clstooltip
Dim PathID As Double
Dim PathName As String
Dim IsSelect As Integer
Dim PathValue As Double
Dim RecDate As String
Dim RecTime As String
Private Sub ApproveFlag_Click()
If SystemOptions.UserInterface = ArabicInterface Then
DcbUser2.BoundText = user_id
ApproveTime.value = Time
ApproveDate.value = Date
End If
End Sub
    Function print_report2(Optional NoteSerial As String)
 On Error Resume Next
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 MySQL = " SELECT     dbo.tblbookingrequest2.SDateH, dbo.tblbookingrequest2.ApproveFlag, dbo.tblbookingrequest2.ApproveDate, dbo.tblbookingrequest2.ApproveTime, "
 MySQL = MySQL & "                     dbo.tblbookingrequest2.GroupName, dbo.tblbookingrequest2.ModelID, dbo.tblbookingrequest2.CreationDate, dbo.tblbookingrequest2.ID,"
 MySQL = MySQL & "                     dbo.tblbookingrequest2.SDate, dbo.tblbookingrequest2.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
 MySQL = MySQL & "                     dbo.tblbookingrequest2.InClientID, TblCustemers_1.CusName, TblCustemers_1.CusNamee, TblCustemers_1.Fullcode, dbo.tblbookingrequest2.OutClientID,"
 MySQL = MySQL & "                     TblCustemers_1.CusName AS OutCusName, TblCustemers_1.CusNamee AS OutCusNameE, TblCustemers_1.Fullcode AS OutFullcode,"
 MySQL = MySQL & "                     dbo.tblbookingrequest2.AirPortID, dbo.TblAirport.Name, dbo.TblAirport.NameE, dbo.tblbookingrequest2.AirLineID, dbo.TblAirlines.Name AS AirLineName,"
 MySQL = MySQL & "                     dbo.TblAirlines.NameE AS AirLineNameE, dbo.tblbookingrequest2.EmpName, dbo.tblbookingrequest2.EmpCode, dbo.tblbookingrequest2.EmpMbile,"
 MySQL = MySQL & "                     dbo.tblbookingrequest2.ArriveDate, dbo.tblbookingrequest2.ArriveTime, dbo.tblbookingrequest2.emp, dbo.tblbookingrequest2.other, dbo.tblbookingrequest2.FlightNo,"
 MySQL = MySQL & "                     dbo.tblbookingrequest2.UserID, TblUsers_1.UserName, dbo.tblbookingrequest2.UserID2, TblUsers_1.UserName AS UserName2, dbo.tblbookingrequest2.ProgrammID,"
 MySQL = MySQL & "                      dbo.TblProgrammTypes.Name AS ProgName, dbo.TblProgrammTypes.NameE AS ProgNameE, dbo.tblbookingrequest2.VehicleNo,"
 MySQL = MySQL & "                     dbo.tblbookingrequest2.VehicleType, dbo.TBLCarTypes.name AS VehicTyname, dbo.TBLCarTypes.namee AS VehicTynameE, dbo.TblFlightDetails2.[Date],"
 MySQL = MySQL & "                     dbo.TblFlightDetails2.[Time], dbo.TblFlightDetails2.Remarks, dbo.TblFlightDetails2.PathID, dbo.TblShrines.Name AS PATHName,"
 MySQL = MySQL & "                     dbo.TblShrines.NameE AS PATHNameE, dbo.tblbookingrequest2.Airlounge, dbo.tblbookingrequest2.Remarks AS HRemarks, dbo.tblbookingrequest2.Mobile2,"
 MySQL = MySQL & "                     dbo.tblbookingrequest2.OrdeNo, dbo.tblbookingrequest2.CompnyIn, dbo.tblbookingrequest2.CompnyOut, dbo.tblbookingrequest2.TotalNew,"
 MySQL = MySQL & "                     dbo.tblbookingrequest2.NoteSerial1, dbo.tblbookingrequest2.NoteSerialOrder, dbo.tblbookingrequest2.ProgValue, dbo.tblbookingrequest2.PathAddValue,"
 MySQL = MySQL & "                     dbo.tblbookingrequest2.Total, dbo.tblbookingrequest2.Discount, dbo.tblbookingrequest2.TypeDiscount, dbo.tblbookingrequest2.NetDis,"
 MySQL = MySQL & "                     dbo.tblbookingrequest2.ProAdd, dbo.tblbookingrequest2.SeasonsID, dbo.TblCompaniesGroup.Name AS SeasonsName,"
 MySQL = MySQL & "                     dbo.TblCompaniesGroup.NameE AS SeasonsNameE, dbo.tblbookingrequest2.Prefix, dbo.tblbookingrequest2.HotelMadinh, dbo.tblbookingrequest2.HotelJaddah,"
 MySQL = MySQL & "                     dbo.tblbookingrequest2.HotelMakh , dbo.tblbookingrequest2.FATYou, dbo.tblbookingrequest2.FATValue, dbo.tblbookingrequest2.TotalValue,  TblCustemers_1.VATNO"
 MySQL = MySQL & " FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.tblbookingrequest2 LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCompaniesGroup ON dbo.tblbookingrequest2.SeasonsID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TBLCarTypes ON dbo.tblbookingrequest2.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblProgrammTypes ON dbo.tblbookingrequest2.ProgrammID = dbo.TblProgrammTypes.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblUsers TblUsers_1 ON dbo.tblbookingrequest2.UserID2 = TblUsers_1.UserID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblUsers TblUsers_2 ON dbo.tblbookingrequest2.UserID = TblUsers_2.UserID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblAirlines ON dbo.tblbookingrequest2.AirLineID = dbo.TblAirlines.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblAirport ON dbo.tblbookingrequest2.AirPortID = dbo.TblAirport.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCustemers TblCustemers_1 ON dbo.tblbookingrequest2.OutClientID = TblCustemers_1.CusID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCustemers TblCustemers_2 ON dbo.tblbookingrequest2.InClientID = TblCustemers_2.CusID ON"
 MySQL = MySQL & "                      dbo.TblBranchesData.branch_id = dbo.tblbookingrequest2.BranchID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblShrines RIGHT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblFlightDetails2 ON dbo.TblShrines.ID = dbo.TblFlightDetails2.PathID ON dbo.tblbookingrequest2.ID = dbo.TblFlightDetails2.HID"
 MySQL = MySQL & "  Where (dbo.tblbookingrequest2.ID = " & val(ID.Text) & ")"

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BookingRequest2.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_BookingRequest2.rpt"
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
        
        xReport.ParameterFields(12).AddCurrentValue TxtCusNo.Text
        
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
  xReport.ParameterFields(12).AddCurrentValue TxtCusNo.Text
    End If
     If SystemOptions.VATNoAccordActivity = False Then
    xReport.ParameterFields(11).AddCurrentValue cCompanyInfo.VATRegNo
    Else
    xReport.ParameterFields(11).AddCurrentValue GetRegVATNo(val(BranchID.BoundText))
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(13).AddCurrentValue WriteNo(val(Me.TxtTotalValue), 0, True)

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
 Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "  «„—  ‘€Ì· —ﬁ„ " & ID.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
tablename = "tblbookingrequest2"
Filedname = "ID"
NoteSerial1 = val(ID.Text)
'NoteSerial = val(ID.Text)
Notevalue = 0
 notytype = 9057
Notevalue = val(txtTotal.Text)
BranchID = val(Me.BranchID.BoundText)
NoteDate = (SDate.value)
 
If Notevalue > 0 Then
                              
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des, dtpFromDateH.value         ',
                                              TXTNoteID.Text = NoteID
                                                    TxtNoteSerial.Text = NoteSerial
                                  '  Else
                                  '               If TxtNoteID.text = "" Or TxtNoteSerial.text = "" Then
                                  '       CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                  '                            TxtNoteID.text = NoteID
                                  '                           TxtNoteSerial.text = NoteSerial
                                  '               Else
                                  '                            Sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                  '                            Sql = Sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                  '                               Sql = Sql & " where NoteID=" & val(TxtNoteID.text)
                                  '                               Cn.Execute Sql
                                        
                                  '              End If
                            
Cn.Execute " update tblbookingrequest2 set UserVouchID=" & user_id & " where ID=" & val(ID.Text) & " "
Me.DcbUserVoucher.BoundText = user_id
CREATE_VOUCHER_GE val(TXTNoteID.Text), BranchID, user_id, NoteDate

rs.Resync adAffectCurrent
 

     End If

End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempCustomerCode As String
    Dim StrTempCustomerCodeInsuranceAccount  As String
    
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
 Dim valuee As Double
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
 LngDevNO = 0
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
    my_branch = BranchID
    valuee = val(txtTotal.Text) + val(TxtFATValue.Text)
      StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.OutClientID.BoundText))
      StrTempDes = " «ﬂÌœ «·ÕÃ“ —ﬁ„  " & TxtNoteSerialOrder.Text
      StrTempDes = StrTempDes & CHR(13) & " «„—  ‘€Ì· —ﬁ„  " & TxtNoteSerial1.Text
      StrTempDes = StrTempDes & CHR(13) & "—ﬁ„ «·»—‰«„Ã ·œÌ «·⁄„Ì· —ﬁ„  " & TxtCusNo.Text
      
      

             LngDevNO = LngDevNO + 1
             
              If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 0, StrTempDes & "      Õ”«» «·⁄„Ì· ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            valuee = val(txtTotal.Text)
          StrTempAccountCode = get_account_code_branch(135, my_branch)
         '      StrTempDes = "«„—  ‘€Ì· —ﬁ„   " & ID.Text

          LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & "    Õ”«» «Ì—«œ«  ⁄„—…  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
         '             StrTempDes = "·«„—  ‘€Ì· —ﬁ„   " & ID.Text
If val(TxtFATValue.Text) > 0 Then
          LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVat.BoundText, val(TxtFATValue.Text), 1, "   ﬁÌ„… «·„÷«›…   " & StrTempDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
End If
ErrTrap:
End Function



Private Sub BtnCreateVou_Click()
   If ChekClodePeriod(SDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox " This is Period is Closed"
              End If
              Exit Sub
  End If
If TxtNoteSerial.Text <> "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
Else
MsgBox "Please Delete JE"
End If
Exit Sub
End If

If val(OutClientID.BoundText) = 0 Or OutClientID.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ Ì—ÃÏ «Œ Ì«— «·⁄„Ì· «Ê·«"
Else
MsgBox "Please Select Customer"
End If
Exit Sub
End If
If TxtNoteSerial.Text = "" Then
    Dim Account_Code_dynamic As String
   Account_Code_dynamic = get_account_code_branch(135, my_branch)

    If Account_Code_dynamic = "NO branch" Then
    If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Else
                MsgBox "Please Create Branch"
        End If
                Exit Sub
            Else

                If Account_Code_dynamic = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «Ì—«œ«  «·⁄„—…", vbCritical
                 Else
                 MsgBox "Please Select Account"
                 End If
                   Exit Sub
                End If
            End If
createVoucher
updateNotesValueAndNobytext (val(TXTNoteID.Text))
MsgBox " „ «‰‘«¡ «·ﬁÌœ"
Else
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
End If
End Sub

Private Sub BtnDeleteVoun_Click()
Dim StrSQL As String
        If ChekClodePeriod(SDate.value) = True Then
           If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «·Õ–›   ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox " This is Period is Closed"
           End If
              Exit Sub
        End If
        
'If CheckTab(val(ID.Text)) = True Then
'If SystemOptions.UserInterface = ArabicInterface Then
'MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–« «·«„— „— »ÿ »ÃœÊ· «· —ÕÌ·"
'Else
'MsgBox "Can not be delete this order is linked to a table deportation"
'End If
'Exit Sub
'End If
 If CheckAtfa(val(ID.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–« «·«„— „— »ÿ »‘«‘… «·«ÿ›«¡"
Else
MsgBox "Can not be delete  "
End If
Exit Sub
End If
       StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
       Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TXTNoteID.Text)
       Cn.Execute StrSQL, , adExecuteNoRecords
       Cn.Execute "Update tblbookingrequest2 set NoteSerial=null,UserVouchID=null ,NoteID=null where id=" & val(ID.Text) & ""
       DcbUserVoucher.BoundText = ""
TxtNoteSerial.Text = ""
TXTNoteID.Text = 0
MsgBox " „ Õ–› «·ﬁÌœ"
rs.Resync adAffectCurrent
End Sub

Private Sub Cmd_Click(Index As Integer)
'    On Error GoTo ErrTrap
    Select Case Index
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            BranchID.BoundText = Current_branch
            ID.Text = CStr(new_id("tblbookingrequest2", "ID", "", True))
           emp.value = True
           Grid.Clear flexClearScrollable, flexClearEverything
           Grid.Rows = Grid.FixedRows + 1
           DcbUser.BoundText = user_id
           DcbUser3.BoundText = user_id
           SeasonsID.BoundText = GetMosim(0)
           SDate_Change
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
           If ChekClodePeriod(SDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox " This is Period is Closed"
              End If
              Exit Sub
              End If
              
If TxtNoteSerial.Text <> "" Then
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
Exit Sub
End If
'If CheckTab(val(ID.Text)) = True Then
'If SystemOptions.UserInterface = ArabicInterface Then
'MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–« «·«„— „— »ÿ »ÃœÊ· «· —ÕÌ·"
'Else
'MsgBox "Can not be edited this order is linked to a table deportation"
'End If
'Exit Sub
'End If

If CheckAtfa(val(ID.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–« «·«„— „— »ÿ »‘«‘… «·«ÿ›«¡"
Else
MsgBox "Can not be edite  "
End If
Exit Sub
End If


            TxtModFlg.Text = "E"
            CalResults
            Grid.Rows = Grid.Rows + 1
            SDate_Change
        Case 2
   If ChekClodePeriod(SDate.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
   End If
If val(OutClientID.BoundText) = 0 Or OutClientID.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄„Ì·"
Else
MsgBox "Please Select Customer"
End If
Exit Sub
End If
   Dim TxtNoteSerial1str As String

    If TxtNoteSerial1.Text = "" Then
     TxtNoteSerial1str = Voucher_coding(val(Me.BranchID.BoundText), SDate.value, 71, 71, , , , , , val(SeasonsID.BoundText))
                If TxtNoteSerial1str = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…  Õ—ﬂ…  ÃœÌœ…  ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If TxtNoteSerial1str = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  «·Õ—ﬂ… ÃœÌœ     ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    End If
                End If
    End If
Dim AccountVATDept As String
If AccountVat.BoundText = "" And True = True And CheckAnyVAT = True Then
MsgBox "Ì—ÃÏ ÷»ÿ «⁄œ«œ  «·ﬁÌ„… «·„÷«›…"
Exit Sub
End If
            SaveData

        Case 3
            Undo

        Case 4
        If ChekClodePeriod(SDate.value) = True Then
           If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox " This is Period is Closed"
           End If
              Exit Sub
        End If
              
If TxtNoteSerial.Text <> "" Then
MsgBox "Ì—ÃÏ Õ–› «·ﬁÌœ «Ê·«"
Exit Sub
End If

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
 If CheckAtfa(val(ID.Text)) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–« «·«„— „— »ÿ »‘«‘… «·«ÿ›«¡"
Else
MsgBox "Can not be delete  "
End If
Exit Sub
End If
'If CheckTab(val(ID.Text)) = True Then
'If SystemOptions.UserInterface = ArabicInterface Then
'MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–« «·«„— „— »ÿ »ÃœÊ· «· —ÕÌ·"
'Else
'MsgBox "Can not be delete this order is linked to a table deportation"
'End If
'Exit Sub
'End If

        
            Del_Action

        Case 5

        Case 6
                Unload Me
         Case 7
                print_report2
         Case 9
            Unload FrmSearch_Hajj
         FrmSearch_Hajj.SendForm = "BookingRequest2"
         FrmSearch_Hajj.show
         
    End Select

    Exit Sub
ErrTrap:
End Sub



Private Sub CmdAttach_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments ID.Text, "20911201602"
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub DcbTypeDis_Click()
'CalResults
If Me.TxtModFlg.Text <> "R" Then
CalResults
End If
Calculte
'DcbTypeDis_Change
End Sub

Private Sub dtpFromDateH_LostFocus()
            VBA.Calendar = vbCalGreg
            SDate.value = ToGregorianDate(dtpFromDateH.value)
            
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
Function GetProgValu() As Double
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
sql = " Select Valuee from TblProgrammTypes where id=" & val(ProgrammID.BoundText) & ""
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetProgValu = IIf(IsNull(Rs4("Valuee").value), 0, Rs4("Valuee").value)
Else
GetProgValu = 0
End If
End Function
Private Sub Fill_Combos()
 Dim Dcombos As ClsDataCombos
  Dim str As String
   Set Dcombos = New ClsDataCombos
   Dcombos.GetBranches BranchID
   Dcombos.GetCompany InClientID, 0, 0
   Dcombos.GetCompany OutClientID, 2, 0
   Dcombos.GetAccountingCodes AccountVat
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
   
    str = "select id , name from tblhotels"
   fill_combo MekkaHotelID, str
    
      str = "select id , name from tblhotels"
   fill_combo JeddahHotelID, str
   
     str = "select id , name from tblhotels"
   fill_combo MadinaHotelID, str
    If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblCompaniesGroup  "
   Else
   str = " select id , nameE from TblCompaniesGroup  "
  End If
  str = str & " where Omra_Hajj=0"
   fill_combo SeasonsID, str
   Dcombos.GetTblCarsDataGroup VehicleType, True
   
   Dcombos.GetUsers Me.DcbUser
   Dcombos.GetUsers Me.DcbUser2
   Dcombos.GetUsers Me.DcbUser3
   Dcombos.GetUsers Me.DcbUserVoucher
   If SystemOptions.UserInterface = ArabicInterface Then
   With DcbTypeDis
   .Clear
   .AddItem "ﬁÌ„…"
   .AddItem "‰”»…"
   End With
   Else
  With DcbTypeDis
   .Clear
   .AddItem "Value"
   .AddItem "Percentage"
   End With
   End If
   ' Dcombos.getCountriesGovernments Me.inCity
End Sub


Private Sub Form_Load()
 '   On Error GoTo ErrTrap
Dim i As Integer
  If SystemOptions.AllowCreateHajomraVoucher = True Then
    BtnCreateVou.Enabled = True
BtnDeleteVoun.Enabled = True
Ele(11).Enabled = True
 Else
  BtnCreateVou.Enabled = False
 BtnDeleteVoun.Enabled = False
 Ele(11).Enabled = False
 End If
 
        Fill_Combos
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  «„—  ‘€Ì·  "
    LogTexte = " Open Window " & " Order Operating "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""
    DcbModelID.Clear
For i = 2015 To 2100
DcbModelID.AddItem i
Next i
    

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
  StrSQL = "SELECT  *  From tblbookingrequest2    "
  Else
 StrSQL = "SELECT  *  From tblbookingrequest2"
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
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «„—  ‘€Ì·   "
    LogTexte = " Exit Window " & "  Order operating "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

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
Dim sql As String
Dim count As Integer
Dim Rate As Double
 Dim str As String
 str = ""
    With Grid

     Select Case .ColKey(Col)
 Case "PathName"
                        StrAccountCode = .ComboData
                        .TextMatrix(Row, .ColIndex("PathID")) = StrAccountCode
                        Grid.Rows = Grid.Rows + 1
  Case "IsPathAdd"
          If .Cell(flexcpChecked, Row, .ColIndex("IsPathAdd")) = flexUnchecked Then
          .TextMatrix(Row, .ColIndex("PathAddValue")) = ""
          End If
     End Select
     If Row = .Rows - 1 And val(.TextMatrix(Row, .ColIndex("PathID"))) <> 0 Then
    
            .Rows = .Rows + 1
        End If
End With
RelainGrid
End Sub
Sub RelainGrid()
Dim i As Integer
Dim Counter As Integer
Dim SumVal As Double
SumVal = 0
Counter = 0
With Grid
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("PathName")) <> "" Then
Counter = Counter + 1
.TextMatrix(i, .ColIndex("Serial")) = Counter
SumVal = SumVal + val(.TextMatrix(i, .ColIndex("PathAddValue")))
End If
Next i
TxtPathAddValue.Text = SumVal * val(VehicleNo.Text)


TxtPathAddValue_Change
End With
End Sub
Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

With Grid

Select Case .ColKey(Col)
     Case "Date"
          .ComboList = ""
     Case "Time"
     .ComboList = ""
     Case "PathName"
    
     Case "Remark"
        .ComboList = ""
        Case "PathAddValue"
           If .Cell(flexcpChecked, .Row, .ColIndex("IsPathAdd")) = flexChecked Then
                    .ComboList = ""
                    Else
                    Cancel = True
                    End If
    End Select
 End With

End Sub



Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Grid
Dim str As String
Select Case Grid.ColKey(Grid.Col)
Case "Date"
           Unload FrmRegesterDateProject
            FrmRegesterDateProject.SendForm = "BookingRequest2"
          FrmRegesterDateProject.show vbModal
Case "Time"
             Unload FrmRegesterDateProject
             FrmRegesterDateProject.SendForm = "BookingRequest2"
             FrmRegesterDateProject.show vbModal
End Select
End With
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
           PathID = val(.TextMatrix(Row, .ColIndex("PathID")))
           PathName = (.TextMatrix(Row, .ColIndex("PathName")))
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
        Case "Date"
        .ColComboList(.ColIndex("Date")) = "..."
         Case "Time"
        .ColComboList(.ColIndex("Time")) = "..."
 
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

Private Sub InClientID_Change()
InClientID_Click (0)
End Sub

Private Sub InClientID_Click(Area As Integer)
   Dim Fullcode As String
    GetCustomersDetail val(InClientID.BoundText), , Fullcode, 1
    Text1.Text = Fullcode
End Sub

Private Sub Model_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.Model.Text, 1)
End Sub

Private Sub OutClientID_Change()
OutClientID_Click (0)
End Sub

Private Sub OutClientID_Click(Area As Integer)
   Dim Fullcode As String
   Dim VATNO As String
    GetCustomersDetail val(OutClientID.BoundText), , Fullcode, 1, , , , , , VATNO
    Text2.Text = Fullcode
    TxtVATNO = VATNO
End Sub

Private Sub SDate_Change()
        dtpFromDateH.value = ToHijriDate(SDate.value)
        ClculteVAT
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim CUSTID As Integer
   If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text1.Text, 2
        InClientID.BoundText = CUSTID
    End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim CUSTID As Integer
 If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text2.Text, 1
        OutClientID.BoundText = CUSTID
    End If
End Sub

Function CalResults()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtProgValue.Text) = 0 Then
TxtProgValue.Text = GetProgValu()
End If
TxtTotalNew.Text = val(TxtProgValue.Text) * val(VehicleNo.Text)
If val(Me.DcbTypeDis.ListIndex) = 1 Then
TxtNetDis.Text = (val(TxtTotalNew.Text) * val(TxtDiscount.Text)) / 100
'TxtNetDis.Text = val(TxtNetDis.Text) * val(VehicleNo.Text)

Else
TxtNetDis.Text = val(TxtDiscount.Text) * val(VehicleNo.Text)
End If
txtTotal.Text = val(TxtTotalNew.Text) - val(TxtNetDis.Text) + val(TxtPathAddValue)

End If
If val(Me.DcbTypeDis.ListIndex) = 1 Then
If SystemOptions.UserInterface = ArabicInterface Then
lbl(25).Caption = "‰”»…"
Else
lbl(25).Caption = "Percentage"
End If
Else
If SystemOptions.UserInterface = ArabicInterface Then
lbl(25).Caption = "ﬁÌ„…"
Else
lbl(25).Caption = "Value"
End If
End If
ClculteVAT
End Function
Sub ClculteVAT()
If Me.TxtModFlg.Text <> "R" Then
Dim Percetage As Double
Dim account As String
PercentgValueAddedAccount_Transec SDate.value, 3, 1, account, Percetage
TxtFATYou.Text = Percetage
AccountVat.BoundText = account
Calculte
End If
End Sub
Sub Calculte()
If Me.TxtModFlg.Text <> "R" Then
If val(TxtFATYou.Text) > 0 Then
TxtFATValue.Text = (val(txtTotal.Text) * val(TxtFATYou.Text)) / 100
Else
TxtFATValue.Text = 0
End If
TxtTotalValue.Text = val(txtTotal.Text) + val(TxtFATValue.Text)
End If
End Sub
Private Sub txtDiscount_Change()
If Me.TxtModFlg.Text <> "R" Then
CalResults
End If

'If Me.TxtModFlg.Text <> "R" Then
'If val(Me.DcbTypeDis.ListIndex) = 1 Then
'TxtNetDis.Text = (val(TxtProAdd.Text) * val(TxtDiscount.Text)) / 100
'Else
'TxtNetDis.Text = val(TxtDiscount.Text)
'End If
'TxtTotal.Text = val(TxtProAdd.Text) - val(TxtNetDis.Text)
'End If
End Sub

Private Sub TxtFATYou_Change()
Calculte
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   «„—  ‘€Ì·"
            Else
                Me.Caption = "Order Operating  Data"
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
                Me.Caption = "»Ì«‰«  «„— ‘€· ( ÃœÌœ )"
            Else
                Me.Caption = "Order Operating Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   «„—  ‘€Ì· ( ÃœÌœ )"
            Else
                Me.Caption = "Order Operarting Data(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            ID.locked = True
            pnlHeader.Enabled = True
            SDate.Enabled = True
            Me.BranchID.Enabled = True
        Case "E"

        If SystemOptions.DateCanNotEdit = True Then
            Me.SDate.Enabled = False
            Else
            Me.SDate.Enabled = True
            End If

                         If SystemOptions.BranchCanNotEdit = True Then
                            Me.BranchID.Enabled = False
                            Else
                              Me.BranchID.Enabled = True
                            End If

      

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«    «„—  ‘€Ì· (  ⁄œÌ· )"
            Else
                Me.Caption = "Order Operatingt Data(Edit)"
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
Function CheckExist(Optional OrdeNo As Double, Optional ByRef idd As Double, Optional SeasonsID As Double) As Boolean
Dim sql As String
Dim Rs1 As ADODB.Recordset
Set Rs1 = New ADODB.Recordset
    sql = "Select * from tblbookingrequest2 where  OrdeNo=" & OrdeNo & " and ID<>" & val(ID.Text) & " "
    Rs1.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
    idd = IIf(IsNull(Rs1("ID").value), 0, Rs1("ID").value)
    CheckExist = True
    Else
    CheckExist = False
    End If
End Function

Public Sub RetriveOrder(Optional Lngid As Double = 0)
Dim sql As String
Dim Rs1 As ADODB.Recordset
Set Rs1 = New ADODB.Recordset
    'On Error GoTo ErrTrap
    If Lngid <> 0 Then
    End If
    '  If Me.TxtModFlg.Text = "R" Then
    'Sql = "Select * from tblbookingrequest where id=" & Lngid
    'Else
    sql = "Select * from tblbookingrequest where id=" & Lngid & " and  StusID=1     "
    'End If
    Rs1.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  ' Grid.Clear flexCleaRS1crollable, flexClearEverything
  '         Grid.Rows = Grid.FixedRows + 1
     If Rs1.RecordCount <= 0 Then

       If Me.TxtModFlg.Text = "R" Then
 Exit Sub
 End If
     Dim str As String
     str = ID.Text
    clear_all Me
    ID.Text = str
      Grid.Clear flexClearScrollable, flexClearEverything
           Grid.Rows = Grid.FixedRows + 1
     Else
      If Me.TxtModFlg.Text = "R" Then
    
    Exit Sub
 End If
 TxtCusNo.Text = IIf(IsNull(Rs1("CusNo").value), "", Trim(Rs1("CusNo").value))
   ' ID.text = IIf(IsNull(RS1("ID").value), "", (RS1("ID").value))
   TxtCompnyIn.Text = IIf(IsNull(Rs1("CompnyIn").value), "", Trim(Rs1("CompnyIn").value))
    TxtCompnyOut.Text = IIf(IsNull(Rs1("CompnyOut").value), "", Trim(Rs1("CompnyOut").value))

   Me.SeasonsID.BoundText = IIf(IsNull(Rs1("SeasonsID").value), "", (Rs1("SeasonsID").value))
    TxtHotelMakh.Text = IIf(IsNull(Rs1("HotelMakh").value), "", (Rs1("HotelMakh").value))
    TxtHotelMadinh.Text = IIf(IsNull(Rs1("HotelMadinh").value), "", (Rs1("HotelMadinh").value))
    TxtHotelJaddah.Text = IIf(IsNull(Rs1("HotelJaddah").value), "", (Rs1("HotelJaddah").value))
    SDate.value = IIf(IsNull(Rs1("Sdate").value), Date, Rs1("Sdate").value)
    BranchID.BoundText = IIf(IsNull(Rs1("BranchID").value), "", Trim(Rs1("BranchID").value))
    InClientID.BoundText = IIf(IsNull(Rs1("InClientID").value), "", Trim(Rs1("InClientID").value))
    OutClientID.BoundText = IIf(IsNull(Rs1("OutClientID").value), "", Trim(Rs1("OutClientID").value))
    AirLineID.BoundText = IIf(IsNull(Rs1("AirLineID").value), "", Trim(Rs1("AirLineID").value))
    AirPortID.BoundText = IIf(IsNull(Rs1("AirPortID").value), "", Trim(Rs1("AirPortID").value))
    emp.value = IIf(IsNull(Rs1("emp").value), False, Trim(Rs1("emp").value))
    other.value = IIf(IsNull(Rs1("other").value), False, Trim(Rs1("other").value))
   ' EmpCode.text = IIf(IsNull(Rs1("EmpCode").value), "", Trim(Rs1("EmpCode").value))
    EmpName.Text = IIf(IsNull(Rs1("EmpName").value), "", Trim(Rs1("EmpName").value))
    EmpMbile.Text = IIf(IsNull(Rs1("EmpMbile").value), "", Trim(Rs1("EmpMbile").value))
    FlightNo.Text = IIf(IsNull(Rs1("FlightNo").value), "", Trim(Rs1("FlightNo").value))
    ArriveDate.value = IIf(IsNull(Rs1("ArriveDate").value), Date, Trim(Rs1("ArriveDate").value))
    ArriveTime.value = IIf(IsNull(Rs1("ArriveTime").value), Date, Trim(Rs1("ArriveTime").value))
    ProgrammID.BoundText = IIf(IsNull(Rs1("ProgrammID").value), "", Trim(Rs1("ProgrammID").value))
    VehicleNo.Text = IIf(IsNull(Rs1("VehicleNo").value), 0, Trim(Rs1("VehicleNo").value))
    Model.Text = IIf(IsNull(Rs1("Model").value), "", Trim(Rs1("Model").value))
    'MekkaHotelID.BoundText = IIf(IsNull(Rs1("MekkaHotelID").value), "", Trim(Rs1("MekkaHotelID").value))
    'MadinaHotelID.BoundText = IIf(IsNull(Rs1("MadinaHotelID").value), "", Trim(Rs1("MadinaHotelID").value))
    'JeddahHotelID.BoundText = IIf(IsNull(Rs1("JeddahHotelID").value), "", Trim(Rs1("JeddahHotelID").value))
    VehicleType.BoundText = IIf(IsNull(Rs1("VehicleType").value), "", Trim(Rs1("VehicleType").value))
    GroupID.BoundText = IIf(IsNull(Rs1("GroupID").value), "", Trim(Rs1("GroupID").value))
    DcbModelID.Text = IIf(IsNull(Rs1("ModelID").value), 2016, Trim(Rs1("ModelID").value))
    TxtGroupName.Text = IIf(IsNull(Rs1("GroupName").value), "", Trim(Rs1("GroupName").value))
    Me.DcbUser.BoundText = IIf(IsNull(Rs1("UserID").value), "", Trim(Rs1("UserID").value))
    Me.DcbUser2.BoundText = IIf(IsNull(Rs1("UserID2").value), "", Trim(Rs1("UserID2").value))
    Me.TxtNoteSerialOrder.Text = IIf(IsNull(Rs1("NoteSerial1").value), "", Trim(Rs1("NoteSerial1").value))
    ApproveDate.value = IIf(IsNull(Rs1("ApproveDate").value), Date, Trim(Rs1("ApproveDate").value))
    If Not (IsNull(Rs1("ApproveFlag").value)) Then
    If Rs1("ApproveFlag").value = True Then
    ApproveFlag.value = vbChecked
    Else
    ApproveFlag.value = vbUnchecked
    End If
    End If
    Dim ContactTime As Date
     If Not IsNull(Rs1("ApproveTime").value) Then
     ContactTime = FormatDateTime(Rs1("ApproveTime").value, vbShortTime)
      Me.ApproveTime.value = ContactTime
    End If

    
    
    Set Rs_Temp = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT     dbo.TblFlightDetails.Remarks, dbo.TblFlightDetails.HID, dbo.TblFlightDetails.ID, dbo.TblFlightDetails.[Date], dbo.TblFlightDetails.[Time], dbo.TblFlightDetails.PathID, "
    StrSQL = StrSQL & "                  dbo.TblShrines.name , dbo.TblShrines.NameE"
    StrSQL = StrSQL & " FROM         dbo.TblFlightDetails LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblShrines ON dbo.TblFlightDetails.PathID = dbo.TblShrines.ID"
    StrSQL = StrSQL & "  where TblFlightDetails.HID = " & val(TxtOreder.Text)
    
    Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
     Rs_Temp.MoveFirst
     With Grid
        .Rows = Rs_Temp.RecordCount + 1
        Dim j As Integer
        For j = 1 To .Rows - 1
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
             '''/////////////
                .TextMatrix(j, .ColIndex("olddate")) = IIf(IsNull(Rs_Temp("date").value), "", Rs_Temp("date").value)
                .TextMatrix(j, .ColIndex("oldtime")) = IIf(IsNull(Rs_Temp("time").value), "", Rs_Temp("time").value)
                  .TextMatrix(j, .ColIndex("oldPathID")) = IIf(IsNull(Rs_Temp("PathID").value), 0, Rs_Temp("PathID").value)
                  If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(j, .ColIndex("oldPathName")) = IIf(IsNull(Rs_Temp("Name").value), "", Rs_Temp("Name").value)
                  Else
                    .TextMatrix(j, .ColIndex("oldPathName")) = IIf(IsNull(Rs_Temp("NameE").value), "", Rs_Temp("NameE").value)
                  End If
                Rs_Temp.MoveNext
         Next
        End With
    End If
 End If
    Exit Sub
ErrTrap:
End Sub
Function CheckAtfa(Optional ID As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from tblbookingrequest2 where ID=" & ID & " and FlgExAcc=1 "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckAtfa = True
Else
CheckAtfa = False
End If

End Function
Function CheckTab(Optional ID As Double) As Boolean
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from TblDeported where OrderID=" & ID & " "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CheckTab = True
Else
CheckTab = False
End If

End Function
Public Sub Retrive(Optional Lngid As Long = 0)
Dim ContactTime As Date
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
           
     TxtCusNo.Text = IIf(IsNull(rs("CusNo").value), "", (rs("CusNo").value))
     TxtFATYou.Text = IIf(IsNull(rs("FATYou").value), 0, (rs("FATYou").value))
     TxtFATValue.Text = IIf(IsNull(rs("FATValue").value), 0, (rs("FATValue").value))
     TxtTotalValue.Text = IIf(IsNull(rs("TotalValue").value), 0, (rs("TotalValue").value))
     Me.AccountVat.BoundText = IIf(IsNull(rs("AccountCodeVat").value), "", (rs("AccountCodeVat").value))
    ID.Text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    Me.TxtHotelMakh.Text = IIf(IsNull(rs("HotelMakh").value), "", Trim(rs("HotelMakh").value))
    Me.TxtHotelJaddah.Text = IIf(IsNull(rs("HotelJaddah").value), "", Trim(rs("HotelJaddah").value))
    Me.TxtHotelMadinh.Text = IIf(IsNull(rs("HotelMadinh").value), "", Trim(rs("HotelMadinh").value))
    Me.TxtNoteSerialOrder.Text = IIf(IsNull(rs("NoteSerialOrder").value), "", Trim(rs("NoteSerialOrder").value))
    Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    Me.SeasonsID.BoundText = IIf(IsNull(rs.Fields("SeasonsID").value), "", rs.Fields("SeasonsID").value)
    Me.DcbUser3.BoundText = IIf(IsNull(rs.Fields("UserID").value), "", rs.Fields("UserID").value)
    DcbUserVoucher.BoundText = IIf(IsNull(rs.Fields("UserVouchID").value), "", rs.Fields("UserVouchID").value)
    TxtNoteSerial.Text = IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value)
    Me.TXTNoteID.Text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
    SDate.value = IIf(IsNull(rs("Sdate").value), Date, rs("Sdate").value)
    dtpFromDateH.value = ToHijriDate(SDate.value)  'IIf(IsNull(rs("SDateH").value), ToHijriDate(Date), rs("SDateH").value)
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
    If Not IsNull(rs("ArriveTime").value) Then
     ContactTime = FormatDateTime(rs("ArriveTime").value, vbShortTime)
      Me.ApproveTime.value = ContactTime
    End If
TxtCompnyIn.Text = IIf(IsNull(rs("CompnyIn").value), "", Trim(rs("CompnyIn").value))
TxtCompnyOut.Text = IIf(IsNull(rs("CompnyOut").value), "", Trim(rs("CompnyOut").value))
    ProgrammID.BoundText = IIf(IsNull(rs("ProgrammID").value), "", Trim(rs("ProgrammID").value))
    VehicleNo.Text = IIf(IsNull(rs("VehicleNo").value), 0, Trim(rs("VehicleNo").value))
    Model.Text = IIf(IsNull(rs("Model").value), "", Trim(rs("Model").value))
    MekkaHotelID.BoundText = IIf(IsNull(rs("MekkaHotelID").value), "", Trim(rs("MekkaHotelID").value))
    MadinaHotelID.BoundText = IIf(IsNull(rs("MadinaHotelID").value), "", Trim(rs("MadinaHotelID").value))
    JeddahHotelID.BoundText = IIf(IsNull(rs("JeddahHotelID").value), "", Trim(rs("JeddahHotelID").value))
    VehicleType.BoundText = IIf(IsNull(rs("VehicleType").value), "", Trim(rs("VehicleType").value))
    'GroupID.BoundText = IIf(IsNull(rs("GroupID").value), "", Trim(rs("GroupID").value))
    DcbModelID.Text = IIf(IsNull(rs("ModelID").value), -1, Trim(rs("ModelID").value))
    TxtGroupName.Text = IIf(IsNull(rs("GroupName").value), "", Trim(rs("GroupName").value))
    Me.DcbUser.BoundText = IIf(IsNull(rs("UserID").value), "", Trim(rs("UserID").value))
    Me.DcbUser2.BoundText = IIf(IsNull(rs("UserID2").value), "", Trim(rs("UserID2").value))
    ApproveDate.value = IIf(IsNull(rs("ApproveDate").value), Date, Trim(rs("ApproveDate").value))
    If Not (IsNull(rs("CompanyID").value)) Then
    If rs("CompanyID").value = True Then
    ApproveFlag.value = vbChecked
    Else
    ApproveFlag.value = vbUnchecked
    End If
    End If
    TxtAirlounge.Text = IIf(IsNull(rs("Airlounge").value), "", Trim(rs("Airlounge").value))
    TXtMobile2.Text = IIf(IsNull(rs("Mobile2").value), "", Trim(rs("Mobile2").value))
    TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", Trim(rs("Remarks").value))
    TxtProgValue.Text = IIf(IsNull(rs("ProgValue").value), "", Trim(rs("ProgValue").value))
    TxtPathAddValue.Text = IIf(IsNull(rs("PathAddValue").value), "", Trim(rs("PathAddValue").value))
    txtTotal.Text = IIf(IsNull(rs("Total").value), "", Trim(rs("Total").value))
    TxtDiscount.Text = IIf(IsNull(rs("Discount").value), "", Trim(rs("Discount").value))
    TxtNetDis.Text = IIf(IsNull(rs("NetDis").value), "", Trim(rs("NetDis").value))
    TxtProAdd.Text = IIf(IsNull(rs("ProAdd").value), "", Trim(rs("ProAdd").value))
    Me.DcbTypeDis.ListIndex = IIf(IsNull(rs("TypeDiscount").value), -1, Trim(rs("TypeDiscount").value))
    
     If Not IsNull(rs("ApproveTime").value) Then
     ContactTime = FormatDateTime(rs("ApproveTime").value, vbShortTime)
      Me.ApproveTime.value = ContactTime
    End If
    TxtOreder.Text = IIf(IsNull(rs("OrdeNo").value), "", Trim(rs("OrdeNo").value))
    TxtOreder2.Text = IIf(IsNull(rs("OrdeNo").value), "", Trim(rs("OrdeNo").value))
    TxtTotalNew = val(TxtProgValue.Text) * val(VehicleNo.Text)
    Set Rs_Temp = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "SELECT     dbo.TblFlightDetails2.Remarks, dbo.TblFlightDetails2.HID, dbo.TblFlightDetails2.ID, dbo.TblFlightDetails2.[Date], dbo.TblFlightDetails2.[Time], dbo.TblFlightDetails2.PathID, "
    StrSQL = StrSQL & "                  dbo.TblShrines.name , dbo.TblShrines.NameE ,dbo.TblFlightDetails2.IsPathAdd,dbo.TblFlightDetails2.PathAddValue"
    StrSQL = StrSQL & " FROM         dbo.TblFlightDetails2 LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblShrines ON dbo.TblFlightDetails2.PathID = dbo.TblShrines.ID"
    StrSQL = StrSQL & "  where TblFlightDetails2.HID = " & val(ID.Text)
    
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
                .TextMatrix(j, .ColIndex("PathAddValue")) = IIf(IsNull(Rs_Temp("PathAddValue").value), "", Rs_Temp("PathAddValue").value)
                .TextMatrix(j, .ColIndex("IsPathAdd")) = IIf(IsNull(Rs_Temp("IsPathAdd").value), 0, Rs_Temp("IsPathAdd").value)
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
                 ''///////////////////////
                  .TextMatrix(j, .ColIndex("olddate")) = IIf(IsNull(Rs_Temp("date").value), "", Rs_Temp("date").value)
                  .TextMatrix(j, .ColIndex("oldtime")) = IIf(IsNull(Rs_Temp("time").value), "", Rs_Temp("time").value)
                  .TextMatrix(j, .ColIndex("oldPathID")) = IIf(IsNull(Rs_Temp("PathID").value), 0, Rs_Temp("PathID").value)
                  .TextMatrix(j, .ColIndex("oldPathAddValue")) = 0
                  .TextMatrix(j, .ColIndex("oldIsPathAdd")) = 0
                
                  If SystemOptions.UserInterface = ArabicInterface Then
                  .TextMatrix(j, .ColIndex("oldPathName")) = IIf(IsNull(Rs_Temp("Name").value), "", Rs_Temp("Name").value)
                  Else
                  .TextMatrix(j, .ColIndex("oldPathName")) = IIf(IsNull(Rs_Temp("NameE").value), "", Rs_Temp("NameE").value)
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


Sub RetriveCompanyCon()
Dim sql As String
Dim TypeDiscount As Integer
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     dbo.TblCompanyContract.ID, dbo.TblCompanyContract.LockedID, dbo.TblCompanyContract.CompID, dbo.TblCompanyContract.FromDate,"
sql = sql & "                       dbo.TblCompanyContract.Todate, dbo.TblCompanyContractDet.Discount, dbo.TblCompanyContractDet.NetDis, dbo.TblCompanyContractDet.TypeDiscount,"
sql = sql & "                       dbo.TblCompanyContractDet.ProjID ,dbo.TblCompanyContractDet.Price"
sql = sql & " FROM         dbo.TblCompanyContract LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCompanyContractDet ON dbo.TblCompanyContract.ID = dbo.TblCompanyContractDet.CompContID"
sql = sql & "  Where (dbo.TblCompanyContract.LockedID = 0)"
sql = sql & "  and  (dbo.TblCompanyContractDet.ProjID = " & val(ProgrammID.BoundText) & ")"
sql = sql & "  and  (dbo.TblCompanyContract.CompID = " & val(OutClientID.BoundText) & ")"
sql = sql & "  and  (dbo.TblCompanyContract.FromDate <= " & SQLDate(SDate.value, True) & ")"
sql = sql & "  and  (dbo.TblCompanyContract.Todate >= " & SQLDate(SDate.value, True) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
TxtProgValue.Text = IIf(IsNull(Rs3("Price").value), 0, Rs3("Price").value)
TypeDiscount = IIf(IsNull(Rs3("TypeDiscount").value), 0, Rs3("TypeDiscount").value)
TypeDiscount = TypeDiscount - 1
DcbTypeDis.ListIndex = TypeDiscount
TxtDiscount.Text = IIf(IsNull(Rs3("Discount").value), 0, Rs3("Discount").value)
TxtDiscount.Text = IIf(IsNull(Rs3("Discount").value), 0, Rs3("Discount").value)
TxtNetDis.Text = IIf(IsNull(Rs3("NetDis").value), 0, Rs3("NetDis").value)

If Me.TxtModFlg.Text <> "R" Then
CalResults
End If

Else
DcbTypeDis.ListIndex = -1
TxtDiscount.Text = 0
TxtNetDis.Text = 0
TxtTotalNew.Text = 0
End If
End Sub

Private Sub TxtNetDis_Change()
Calculte
End Sub

Private Sub TxtNoteSerialOrder_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If val(SeasonsID.BoundText) <> 0 Then
TxtOreder.Text = GetIDOrder(val(TxtNoteSerialOrder.Text), val(SeasonsID.BoundText))
Booking
Else
MsgBox "Ì—ÃÏ  ÕœÌœ «·„Ê”„ «Ê·«"
SeasonsID.SetFocus
End If
End If
End Sub
Public Sub Booking()
Dim idd As Double
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then

If CheckExist(val(TxtOreder.Text), idd) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "—ﬁ„  «ﬂÌœ «·ÕÃ“ „ÊÃÊœ „”»ﬁ« ›Ì «„—  ‘€Ì· —ﬁ„" & " " & ID
Else
MsgBox " This number is in the operating No." & " " & ID
End If
     Dim str As String
     str = ID.Text
    clear_all Me
    ID.Text = str
      Grid.Clear flexClearScrollable, flexClearEverything
           Grid.Rows = Grid.FixedRows + 1
Exit Sub
End If
RetriveOrder val(TxtOreder.Text)

TxtTotalNew = val(TxtProgValue.Text) * val(VehicleNo.Text)
'CalResults
If Me.TxtModFlg.Text <> "R" Then
CalResults
End If

RetriveCompanyCon

End If
End Sub

Private Sub TxtNoteSerialOrder_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.Text <> "R" Then
If KeyCode = vbKeyF3 Then
    Unload FrmSearch_Hajj
         FrmSearch_Hajj.SendForm = "BookingRequest11"
         FrmSearch_Hajj.show
End If
End If
End Sub

Private Sub TxtOreder_Change()
'If Me.TxtModFlg.Text = "R" Then
'RetriveOrder val(TxtOreder.Text)
'End If
End Sub



Private Sub TxtPathAddValue_Change()
If Me.TxtModFlg.Text <> "R" Then
CalResults
'TxtProAdd.Text = val(Me.TxtTotalNew.Text)
'TxtTotal.Text = val(TxtProAdd.Text) - val(TxtNetDis.Text)
'DcbTypeDis_Change
End If
End Sub

Private Sub TxtProgValue_Change()
If Me.TxtModFlg.Text <> "R" Then
'TxtProAdd.Text = (val(Me.TxtProgValue.Text) + val(TxtPathAddValue.Text)) * val(VehicleNo.Text)
'TxtTotal.Text = val(TxtProAdd.Text) - val(TxtNetDis.Text)
CalResults

'DcbTypeDis_Change
End If
End Sub

Private Sub TxtTotal_Change()
Calculte
End Sub

Private Sub TxtTotalNew_Change()
Calculte
End Sub

Private Sub VehicleNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.VehicleNo.Text, 1)
End Sub

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
                ID.Text = CStr(new_id("tblbookingrequest2", "ID", "", True))
           Case "E"
               Cn.Execute "Update tblbookingrequest set UseFlag=Null where ID=" & val(TxtOreder2.Text) & "  "
                StrSQL = "delete From TblFlightDetails2 where  HID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 
           End Select
        rs("ID").value = val(ID.Text)
        If TxtNoteSerial1.Text = "" Then
              TxtNoteSerial1.Text = Voucher_coding(val(Me.BranchID.BoundText), SDate.value, 71, 71, , , , , , val(SeasonsID.BoundText))
        End If
        rs("CusNo").value = TxtCusNo.Text
        rs("FATYou").value = val(TxtFATYou.Text)
        rs("FATValue").value = val(TxtFATValue.Text)
        rs("TotalValue").value = val(TxtTotalValue.Text)
        rs("AccountCodeVat").value = Me.AccountVat.BoundText
        rs("NoteSerial1").value = IIf(Me.TxtNoteSerial1 <> "", val(TxtNoteSerial1.Text), Null)
        rs("NoteSerialOrder").value = IIf(Me.TxtNoteSerialOrder <> "", val(TxtNoteSerialOrder.Text), Null)
        rs("SeasonsID").value = val(SeasonsID.BoundText)
        rs("SDate").value = SDate.value
        rs("HotelMakh").value = TxtHotelMakh.Text
        rs("HotelJaddah").value = TxtHotelJaddah.Text
        rs("HotelMadinh").value = TxtHotelMadinh.Text
        rs("SDateH").value = dtpFromDateH.value
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
        rs("ProgrammID").value = IIf(ProgrammID.BoundText = "", Null, (ProgrammID.BoundText))
        rs("VehicleNo").value = IIf(VehicleNo.Text = "", 0, val(VehicleNo.Text))
        rs("Model").value = IIf(Model.Text = "", 0, Model.Text)
        rs("MekkaHotelID").value = IIf(MekkaHotelID.BoundText = "", Null, (MekkaHotelID.BoundText))
        rs("JeddahHotelID").value = IIf(JeddahHotelID.BoundText = "", Null, (JeddahHotelID.BoundText))
        rs("MadinaHotelID").value = IIf(MadinaHotelID.BoundText = "", Null, (MadinaHotelID.BoundText))
        rs("VehicleType").value = IIf(VehicleType.BoundText = "", Null, (VehicleType.BoundText))
        rs("CompnyIn").value = TxtCompnyIn.Text
       rs("CompnyOut").value = TxtCompnyOut.Text
      '  rs("EmpCode").value = IIf(EmpCode.text = "", Null, (EmpCode.text))
        rs("EmpName").value = IIf(EmpName.Text = "", Null, (EmpName.Text))
        rs("EmpMbile").value = IIf(EmpMbile.Text = "", Null, (EmpMbile.Text))
        rs("VehicleType").value = IIf(VehicleType.BoundText = "", Null, (VehicleType.BoundText))
        rs("emp").value = emp.value
        rs("other").value = other.value
        rs("FlightNo").value = FlightNo.Text
        rs("creationdate").value = Date
        rs("creationuserID").value = user_id
        rs("UserID").value = IIf(DcbUser3.BoundText = "", Null, val(DcbUser3.BoundText))
        rs("GroupID").value = IIf(GroupID.BoundText = "", Null, (GroupID.BoundText))
        rs("ModelID").value = IIf(val(DcbModelID.Text) = -1, Null, val(DcbModelID.Text))
        rs("CompanyID").value = IIf(CompanyID.BoundText = "", Null, (CompanyID.BoundText))
        rs("UserID").value = IIf(Me.DcbUser.BoundText = "", Null, val(DcbUser.BoundText))
        rs("UserID2").value = IIf(DcbUser2.BoundText = "", Null, val(DcbUser2.BoundText))
        rs("ApproveTime").value = FormatDateTime(ApproveTime.value, vbShortTime)
        rs("GroupName").value = IIf(Me.TxtGroupName.Text = "", Null, (TxtGroupName.Text))
        rs("ApproveDate").value = FormatDateTime(ApproveDate.value, vbShortTime)
        If ApproveFlag.value = vbChecked Then
        rs("CompanyID").value = 1
        Else
        rs("CompanyID").value = 0
        End If
        rs("Airlounge").value = TxtAirlounge.Text
        rs("Mobile2").value = TXtMobile2.Text
        rs("Remarks").value = TxtRemarks.Text
        rs("OrdeNo").value = val(TxtOreder.Text)
        rs("ProgValue").value = val(TxtProgValue.Text)
        rs("PathAddValue").value = val(TxtPathAddValue.Text)
        rs("Total").value = val(txtTotal.Text)
        rs("Discount").value = val(TxtDiscount.Text)
        rs("NetDis").value = val(TxtNetDis.Text)
        rs("ProAdd").value = val(TxtProAdd.Text)
        rs("TypeDiscount").value = val(Me.DcbTypeDis.ListIndex)
        rs.update
        Cn.Execute "Update tblbookingrequest set UseFlag=1 where ID=" & val(TxtOreder.Text) & "  "
       Dim Rs_Temp As ADODB.Recordset
        Set Rs_Temp = New ADODB.Recordset
        StrSQL = " select * from TblFlightDetails2  where 1 = -1 "
        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        With Grid
        Dim j As Integer
        For j = 1 To Grid.Rows - 1
           If val(.TextMatrix(j, .ColIndex("PathID"))) <> 0 Then
                    Rs_Temp.AddNew
                    Rs_Temp("ID").value = CStr(new_id("TblFlightDetails2", "ID", "", True))
                    Rs_Temp("HID").value = val(ID.Text)
                    Rs_Temp("PathID").value = val(.TextMatrix(j, .ColIndex("PathID")))
                    Rs_Temp("PathAddValue").value = val(.TextMatrix(j, .ColIndex("PathAddValue")))
                    If .Cell(flexcpChecked, j, .ColIndex("IsPathAdd")) = flexChecked Then
                    Rs_Temp("IsPathAdd") = 1
                    Else
                    Rs_Temp("IsPathAdd") = 0
                    End If
                    Rs_Temp("Date").value = IIf(.TextMatrix(j, .ColIndex("Date")) = "", Date, .TextMatrix(j, .ColIndex("Date")))
                    Rs_Temp("Time").value = IIf(.TextMatrix(j, .ColIndex("Time")) = "", "", .TextMatrix(j, .ColIndex("Time")))
                    Rs_Temp("Remarks").value = .TextMatrix(j, .ColIndex("Remark"))
                    Rs_Temp("creationdate").value = Date
                    Rs_Temp("creationuserID").value = user_id
                    Rs_Temp("oldtime").value = IIf(.TextMatrix(j, .ColIndex("oldtime")) = "", "", .TextMatrix(j, .ColIndex("oldtime")))
                    Rs_Temp("olddate").value = IIf(.TextMatrix(j, .ColIndex("olddate")) = "", Date, .TextMatrix(j, .ColIndex("olddate")))
                    Rs_Temp("oldPathID") = val(.TextMatrix(j, .ColIndex("oldPathID")))
                    Rs_Temp("oldPathAddValue").value = val(.TextMatrix(j, .ColIndex("oldPathAddValue")))
                     If .Cell(flexcpChecked, j, .ColIndex("oldIsPathAdd")) = flexChecked Then
                    Rs_Temp("oldIsPathAdd") = 1
                    Else
                    Rs_Temp("oldIsPathAdd") = 0
                    End If
                    Rs_Temp.update
                 End If
           Next
        End With
            
        Set Rs_Temp = New ADODB.Recordset
        StrSQL = " select * from TblboKRegister  where 1 = -1 "
        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        With Grid
        Dim str As String
        For j = 1 To Grid.Rows - 1
        str = ""
           If val(.TextMatrix(j, .ColIndex("PathID"))) <> 0 Then
             If val(.TextMatrix(j, .ColIndex("PathID"))) <> val(.TextMatrix(j, .ColIndex("oldPathID"))) Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                 str = str & " „  ⁄œÌ· „‰ «·„”«—" & " " & .TextMatrix(j, .ColIndex("oldPathName"))
                 str = str & "«·Ï  «·„”«—" & " " & .TextMatrix(j, .ColIndex("PathName"))
                 Else
                 str = str & "Update From Path" & " " & .TextMatrix(j, .ColIndex("oldPathName"))
                 str = str & "To Path" & " " & .TextMatrix(j, .ColIndex("PathName"))
                 str = str & CHR(13)
                 End If
                 End If
                 
                 If .TextMatrix(j, .ColIndex("Date")) <> .TextMatrix(j, .ColIndex("olddate")) Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                 str = str & " „  ⁄œÌ· „‰  «—ÌŒ" & " " & .TextMatrix(j, .ColIndex("olddate"))
                 str = str & "«·Ï   «—ÌŒ" & " " & .TextMatrix(j, .ColIndex("Date"))
                 Else
                 str = str & "Update From Date" & " " & .TextMatrix(j, .ColIndex("olddate"))
                 str = str & "To Date" & " " & .TextMatrix(j, .ColIndex("Date"))
                 str = str & CHR(13)
                 End If
                 End If
                  If (.TextMatrix(j, .ColIndex("oldtime"))) <> (.TextMatrix(j, .ColIndex("time"))) Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                 str = str & " „  ⁄œÌ· „‰ «·Êﬁ " & " " & .TextMatrix(j, .ColIndex("oldtime"))
                 str = str & "«·Ï  «·Êﬁ " & " " & .TextMatrix(j, .ColIndex("time"))
                 Else
                 str = str & "Update From Time " & " " & .TextMatrix(j, .ColIndex("oldtime"))
                 str = str & "To Time" & " " & .TextMatrix(j, .ColIndex("time"))
                 str = str & CHR(13)
                 End If
                 End If
                 
               
                   If val(.TextMatrix(j, .ColIndex("PathAddValue"))) <> val(.TextMatrix(j, .ColIndex("oldPathAddValue"))) Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                 str = str & " „  ⁄œÌ· ﬁÌ„… «·„”«—" & " " & .TextMatrix(j, .ColIndex("oldPathAddValue"))
                 str = str & "«·Ï  «·ﬁÌ„…" & " " & .TextMatrix(j, .ColIndex("PathAddValue"))
                 Else
                 str = str & "Update From Value " & " " & .TextMatrix(j, .ColIndex("oldPathAddValue"))
                 str = str & "To Value" & " " & .TextMatrix(j, .ColIndex("PathAddValue"))
                 str = str & CHR(13)
                 End If
                 End If
                 If str <> "" Then
                    Rs_Temp.AddNew
                    Rs_Temp("ID") = CStr(new_id("TblboKRegister", "ID", "", True))
                    Rs_Temp("OderNo") = val(ID.Text)
                    Rs_Temp("RecTime") = Time
                    Rs_Temp("RecDate") = Date
                    Rs_Temp("Remarks").value = str
                    Rs_Temp("UserName").value = Me.DcbUser3.Text
                     Rs_Temp.update
                  End If
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
                    Msg = "  „ Õ›Ÿ »Ì«‰«   " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
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
        Retrive val(ID.Text)
    End If

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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Grid.Rows = Grid.FixedRows
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)


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
 
    ' On Error GoTo ErrTrap
            
        If ID.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«   —ﬁ„ " & CHR(13)
        Msg = Msg + (ID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        Msg = "Confirm Delete  " & CHR(13)
        Msg = Msg + (ID.Text) & CHR(13)
        Msg = Msg + "  Are you sure you want to delete ?"
        End If
        
        
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                  Cn.Execute "Update tblbookingrequest set UseFlag=Null where ID=" & val(TxtOreder.Text) & "  "
                StrSQL = "delete From TblFlightDetails2 where  HID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
                StrSQL = "delete From tblbookingrequest2 where  ID =" & val(ID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 
                 rs.MoveFirst
                    
                   StrSQL = "SELECT  *  From tblbookingrequest2 "
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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ…  "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    
    rs.CancelUpdate
    'End If

End Sub



Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰« «„—  ‘€Ì·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «„—  ‘€Ì·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «„—  ‘€Ì· «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «„—  ‘€Ì·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "«„—  ‘€Ì·" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    
    With TTP
        .Create Me.hwnd, "»Ì«‰«  «„—  ‘€Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub
