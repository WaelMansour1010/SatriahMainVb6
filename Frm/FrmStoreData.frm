VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmStoreData 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·„Œ«“‰"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   HelpContextID   =   180
   Icon            =   "FrmStoreData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   8535
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   150
      Visible         =   0   'False
      Width           =   855
   End
   Begin ImpulseButton.ISButton XPBtnMove 
      Height          =   345
      Index           =   0
      Left            =   1125
      TabIndex        =   1
      Top             =   150
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
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
      ButtonImage     =   "FrmStoreData.frx":038A
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
      Left            =   60
      TabIndex        =   2
      Top             =   150
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
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
      ButtonImage     =   "FrmStoreData.frx":0724
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
      Left            =   1650
      TabIndex        =   3
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
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
      ButtonImage     =   "FrmStoreData.frx":0ABE
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
      Left            =   585
      TabIndex        =   4
      Top             =   150
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   609
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
      ButtonImage     =   "FrmStoreData.frx":0E58
      ColorHighlight  =   4194304
      ColorHoverText  =   16777215
      ColorShadow     =   -2147483631
      ColorOutline    =   -2147483631
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16777215
      ColorTextShadow =   16777215
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   7740
      TabIndex        =   6
      Top             =   8205
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
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
      Left            =   6885
      TabIndex        =   7
      Top             =   8205
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
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
      Left            =   6030
      TabIndex        =   8
      Top             =   8205
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
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
      Left            =   5295
      TabIndex        =   9
      Top             =   8205
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
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
      Left            =   4560
      TabIndex        =   10
      Top             =   8205
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
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
      Left            =   60
      TabIndex        =   11
      Top             =   8205
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
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
      Height          =   345
      Left            =   915
      TabIndex        =   12
      Top             =   8220
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
      ButtonStyle     =   1
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
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   7050
      Left            =   0
      TabIndex        =   17
      Top             =   600
      Width           =   8505
      _cx             =   15002
      _cy             =   12435
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   128
      FrontTabColor   =   14871017
      BackTabColor    =   8454143
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "»Ì‰«‰«  «·„Œ«“‰|«·„Ê«ﬁ⁄|«··«∆ÕÂ «·œ«Œ·Ì…|«·„” Œœ„Ì‰|«·Õ”«»«  «·›—⁄Ì…"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   6
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   0   'False
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   1
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6960
         Left            =   9150
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   45
         Width           =   7200
         _cx             =   12700
         _cy             =   12277
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
         GridRows        =   10
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
         Begin VB.TextBox TxtLocation 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   3015
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   4035
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   7185
            _cx             =   12674
            _cy             =   7117
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
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmStoreData.frx":11F2
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
            Height          =   315
            Index           =   20
            Left            =   1785
            TabIndex        =   42
            Top             =   240
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "≈÷«›…"
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
            ButtonImage     =   "FrmStoreData.frx":1248
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   315
            Index           =   21
            Left            =   720
            TabIndex        =   43
            Top             =   240
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
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
            ButtonImage     =   "FrmStoreData.frx":15E2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·„Êﬁ⁄"
            Height          =   315
            Index           =   15
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   1455
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6960
         Index           =   2
         Left            =   45
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   45
         Width           =   7200
         _cx             =   12700
         _cy             =   12277
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
         Begin VB.CheckBox chkIsNotCreateEntry 
            Alignment       =   1  'Right Justify
            Caption         =   "·« Ì‰‘√ ﬁÌœ «·«‰ «Ã"
            Height          =   255
            Left            =   4950
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   6600
            Width           =   1995
         End
         Begin VB.CheckBox chkIsLab 
            Alignment       =   1  'Right Justify
            Caption         =   "„Œ“‰ „⁄„·"
            Height          =   255
            Left            =   2490
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   6630
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox XPMTxtRemark 
            Alignment       =   1  'Right Justify
            Height          =   645
            Left            =   480
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   5850
            Width           =   4065
         End
         Begin VB.CheckBox Chk1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   195
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   7680
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox XPTxtStoreName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   450
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   870
            Width           =   4095
         End
         Begin VB.TextBox XPTxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3450
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox XPTxtStoreAddress 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   450
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1590
            Width           =   4095
         End
         Begin VB.TextBox XPTxtStorePhone 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   450
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1950
            Width           =   4095
         End
         Begin VB.TextBox TXTCode 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   450
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox XPTxtStoreNamee 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   450
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1200
            Width           =   4095
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   450
            TabIndex        =   29
            Top             =   2280
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   450
            TabIndex        =   30
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo SalesPersonid 
            Height          =   315
            Left            =   480
            TabIndex        =   59
            Top             =   4620
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo PurchasePersonid 
            Height          =   315
            Left            =   480
            TabIndex        =   61
            Top             =   4980
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Account_code1 
            Height          =   315
            Index           =   1
            Left            =   450
            TabIndex        =   63
            Top             =   2760
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Account_code3 
            Height          =   315
            Index           =   1
            Left            =   450
            TabIndex        =   65
            Top             =   3240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Account_code2 
            Height          =   315
            Index           =   1
            Left            =   450
            TabIndex        =   67
            Top             =   3660
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Account_code4 
            Height          =   315
            Index           =   1
            Left            =   450
            TabIndex        =   69
            Top             =   4200
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   480
            TabIndex        =   81
            Top             =   5340
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’‰œÊﬁ"
            Height          =   315
            Index           =   23
            Left            =   4650
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   5430
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» Âœ«Ì« Ê⁄Ì‰«  «·—∆Ì”Ì"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   22
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   4230
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» ›ﬁœ Ê ·› «·—∆Ì”Ì"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   21
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   3720
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«»  «· ”ÊÌ«  «·Ã—œÌ… «·—∆Ì”Ì"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   20
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   3240
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «·„Œ“Ê‰ «·—∆Ì”Ì"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   19
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   2880
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰œÊ» «·„‘ —Ì« "
            Height          =   315
            Index           =   18
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   5100
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰œÊ» «·„»Ì⁄« "
            Height          =   315
            Index           =   17
            Left            =   4650
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   4650
            Width           =   2415
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   4560
            TabIndex        =   48
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„Œ“‰"
            Height          =   315
            Index           =   0
            Left            =   1860
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   150
            Width           =   1215
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰ ⁄—»Ì"
            Height          =   315
            Index           =   1
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   870
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   315
            Index           =   3
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   5970
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Â« › «·„Œ“‰"
            Height          =   315
            Index           =   5
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   2040
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄‰Ê«‰ «·„Œ“‰"
            Height          =   315
            Index           =   6
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   1590
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«„Ì‰ «·„” Êœ⁄"
            Height          =   315
            Index           =   7
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   2400
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—»ÿ «·„Œ“‰ »«·„Ã„Ê⁄« "
            Height          =   315
            Index           =   12
            Left            =   5070
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   7920
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   4620
            TabIndex        =   32
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰ «‰Ã·Ì“Ì"
            Height          =   315
            Index           =   13
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1200
            Width           =   2415
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   6960
         Left            =   9450
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   45
         Width           =   7200
         _cx             =   12700
         _cy             =   12277
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
         GridRows        =   10
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
         Begin VB.Frame Frame2 
            Height          =   5010
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   0
            Width           =   7335
            Begin VB.TextBox TxtScreenDesc 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3165
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   46
               Top             =   600
               Width           =   7155
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«··«∆ÕÂ «·œ«Œ·Ì…"
               Height          =   315
               Index           =   14
               Left            =   4800
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   240
               Width           =   2415
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6960
         Index           =   3
         Left            =   9750
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   45
         Width           =   7200
         _cx             =   12700
         _cy             =   12277
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
         Begin VB.Frame Frame10 
            Caption         =   "«”„«¡ «·„” Œœ„Ì‰"
            Height          =   6570
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   210
            Width           =   7110
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   4020
               Left            =   120
               TabIndex        =   51
               Top             =   360
               Width           =   7080
               _cx             =   12488
               _cy             =   7091
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
               Rows            =   1
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmStoreData.frx":1B7C
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
            Begin ALLButtonS.ALLButton CmdRemove 
               Height          =   375
               Left            =   240
               TabIndex        =   52
               Tag             =   "Delete Row"
               Top             =   4560
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "Õ–› ”ÿ—"
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
               COLTYPE         =   2
               FOCUSR          =   -1  'True
               BCOL            =   0
               BCOLO           =   0
               FCOL            =   255
               FCOLO           =   255
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmStoreData.frx":1D3A
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
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Height          =   675
            Index           =   4
            Left            =   645
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   5685
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   300
            Index           =   8
            Left            =   945
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   4260
            Width           =   180
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ⁄·Ìﬁ:"
            Height          =   195
            Index           =   16
            Left            =   945
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Tag             =   "22"
            Top             =   330
            Width           =   390
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   6960
         Left            =   10050
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   45
         Width           =   7200
         _cx             =   12700
         _cy             =   12277
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
         GridRows        =   10
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
         Begin MSDataListLib.DataCombo Account_code1 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   72
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Account_code2 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   73
            Top             =   1290
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Account_code4 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   74
            Top             =   1740
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Account_code3 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   75
            Top             =   870
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õ”«» Âœ«Ì« Ê⁄Ì‰« "
            Height          =   315
            Index           =   11
            Left            =   4110
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   1860
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õ”«»  «· ”ÊÌ«  «·Ã—œÌ…"
            Height          =   315
            Index           =   10
            Left            =   4110
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õ”«» ›ﬁœ Ê ·›"
            Height          =   315
            Index           =   9
            Left            =   4110
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   1350
            Width           =   2415
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õ”«» «·„Œ“Ê‰"
            Height          =   315
            Index           =   8
            Left            =   4110
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   300
            Width           =   2415
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   5
      Left            =   3720
      TabIndex        =   56
      Top             =   8190
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
      ButtonStyle     =   1
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
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   2760
      TabIndex        =   57
      Top             =   8190
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·ﬂ·"
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
      Index           =   8
      Left            =   1800
      TabIndex        =   58
      Top             =   8190
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â "
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
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   2
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   7770
      Width           =   1155
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   4
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   7770
      Width           =   1035
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   7770
      Width           =   945
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   7770
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   " »Ì«‰«  «·„Œ«“‰"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   8475
   End
End
Attribute VB_Name = "FrmStoreData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim Dcombos As ClsDataCombos
Public bo  As Boolean
Public all  As Boolean
Function print_report(Optional NoteSerial As String, Optional X As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    If all = True Then
  MySQL = "  SELECT     dbo.TblStore.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreAdress, dbo.TblStore.StorePhone, dbo.TblStore.Remarks, dbo.TblStore.Account_Code,"
  MySQL = MySQL & "                    dbo.TblStore.Account_Code1, dbo.TblStore.Account_Code2, dbo.TblStore.Account_Code3, dbo.TblStore.linked, dbo.TblStore.Code, dbo.TblStore.StoreNamee,"
 MySQL = MySQL & "                     dbo.TblStore.ParetnAccount, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
 MySQL = MySQL & "                     dbo.TblBranchesData.branch_id , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
MySQL = MySQL & " FROM         dbo.TblStore INNER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblStore.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblStore.Emp_ID = dbo.TblEmployee.Emp_ID"
'MySQL = MySQL & "  Where (dbo.TblCardAuthorizationReform.id =  " & val(XPTxtStoreID.text) & ")"
Else
MySQL = " SELECT     dbo.TblStore.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreAdress, dbo.TblStore.StorePhone, dbo.TblStore.Remarks, dbo.TblStore.Account_Code, "
MySQL = MySQL & "                      dbo.TblStore.Account_Code1, dbo.TblStore.Account_Code2, dbo.TblStore.Account_Code3, dbo.TblStore.linked, dbo.TblStore.Code, dbo.TblStore.StoreNamee,"
MySQL = MySQL & "                      dbo.TblStore.ParetnAccount, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_id , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
MySQL = MySQL & " FROM         dbo.TblStore INNER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblStore.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblStore.Emp_ID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & " Where (dbo.TblStore.StoreId =" & val(XPTxtStoreID.text) & ")"
End If

If all = True Then

  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repallstores.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repallstores.rpt"
 End If
  
Else
 If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repstores.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repstores.rpt"
 End If
End If
  If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation  RepCardAutintcationShow
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
       ' xReport.ParameterFields(15).AddCurrentValue Me.DcboFitter.text
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
    
      '  xReport.ParameterFields(15).AddCurrentValue Me.DcboFitter.text
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  Dim total As String
   Dim dif As String
  Dim totl As Double
  'totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
  'total = totl
  'dif = val(totl) - val(TxtAmoutAccept)
  ' xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
  '    xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
  '      xReport.ParameterFields(14).AddCurrentValue total
  '      xReport.ParameterFields(15).AddCurrentValue dif
       
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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

Private Sub RemoveGridRow()

    With Me.Grid

        If .Row <= 0 Then
            Exit Sub
        End If
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("Location")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

    With Me.VSFlexGrid1

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("id")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
  
            End If

        Next i
   
    End With
    
End Sub

Function addrow()
    Dim Msg As String

    If TxtLocation.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»       «œŒ«· «”„ «·„Êﬁ⁄  ...!!!"
        Else
            Msg = "must Enter Location  Name ...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

        TxtLocation.SetFocus
        Exit Function
    End If

    With Grid
 
        .rows = .rows + 1
 
        .TextMatrix(.rows - 1, .ColIndex("Location")) = Me.TxtLocation.text
  
        .AutoSize 0, .Cols - 1, False
 
    End With
 
    TxtLocation.text = ""
    ReLineGrid

End Function

Private Sub Account_code1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.mIndex = Index
        Account_search.case_id = 7897278
    End If
End Sub


Private Sub Account_code2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.mIndex = Index
        Account_search.case_id = 7897279
    End If
End Sub


Private Sub Account_code3_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.mIndex = Index
        Account_search.case_id = 7897280
    End If
End Sub


Private Sub Account_code4_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.mIndex = Index
        Account_search.case_id = 7897281
    End If
End Sub








Private Sub Cmd_Click(Index As Integer)
'    On Error GoTo ErrTrap
    Dim Msg As String

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            XPTxtStoreID.text = CStr(new_id("TblStore", "StoreID", "", True))
      C1Tab1.CurrTab = 0
      
       '     XPTxtStoreName.SetFocus
            Me.Chk1.value = vbChecked
            dcBranch.BoundText = branch_id
VSFlexGrid1.rows = 1
            Account_code1(1).Enabled = True
            Account_code2(1).Enabled = True
            Account_code3(1).Enabled = True
            Account_code4(1).Enabled = True
            Dcbranch_Click 0
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
             VSFlexGrid1.rows = VSFlexGrid1.rows + 1
            VSFlexGrid1.Enabled = True
            
            TxtModFlg.text = "E"
            CuurentLogdata

        Case 2
        Dim StrVacName As String
          StrVacName = IsRecExist("TblStore", "Code", Trim(TXTCode.text), "Code", "StoreID<>'" & Trim(XPTxtStoreID.text) & "'")

    If StrVacName <> "" Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·ﬂÊœ „‰ ﬁ»·"
   
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
        TXTCode.SetFocus
    
        Exit Sub

    End If
        C1Tab1.CurrTab = 0
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
 
            Del_Company

        Case 5

'bo = True
'wael
 Load FrmStoreSearch
 FrmStoreSearch.show



        Case 6
            Unload Me
            Case 7
            all = True
             If val(Me.XPTxtStoreID.text) <> 0 Then
                print_report val(Me.XPTxtStoreID.text), 1
        
        
            End If
            Case 8
            all = False
             If val(Me.XPTxtStoreID.text) <> 0 Then
                print_report val(Me.XPTxtStoreID.text), 1
        
        
            End If

        Case 20
            addrow

        Case 21
            RemoveGridRow
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub DcboEmp_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetEmployees Me.DcboEmp
    End If

End Sub





Private Sub Dcbranch_Click(Area As Integer)
   
        Account_code1(1).BoundText = get_account_code_branch(0, dcBranch.BoundText)
  
    
 

        Account_code3(1).BoundText = get_account_code_branch(11, dcBranch.BoundText)

        


        Dim currentidAccount As Integer
        If SystemOptions.eachStoreHaveLossAccount = True Then
              currentidAccount = 10
        Else
              currentidAccount = 75
        End If
        Account_code2(1).BoundText = get_account_code_branch(currentidAccount, dcBranch.BoundText)



      If SystemOptions.eachStoreHaveGiftAccount = True Then
            currentidAccount = 17
      Else
            currentidAccount = 76
      End If
    
        Account_code4(1).BoundText = get_account_code_branch(currentidAccount, dcBranch.BoundText)



   
   
    

End Sub

Private Sub Form_Activate()
    'XPTxtStoreID.SetFocus
End Sub

Private Sub Form_Load()
     On Error GoTo ErrTrap
    Dim My_SQL As String
    ScreenNameArabic = "  »Ì«‰«  «·„Œ«“‰ "
    ScreenNameEnglish = " Store Data  "

    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    fill_combo dcBranch, My_SQL

    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
    End If


    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    Resize_Form Me
    Set rs = New ADODB.Recordset
    rs.Open "[TblStore]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"
    AddTip

    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcboEmp
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetAccountingCodes Me.Account_code1(0)
    Dcombos.GetAccountingCodes Me.Account_code2(0)
    Dcombos.GetAccountingCodes Me.Account_code3(0)
    Dcombos.GetAccountingCodes Me.Account_code4(0)
    
    
    Dcombos.GetAccountingCodes Me.Account_code1(1), False, False
    Dcombos.GetAccountingCodes Me.Account_code2(1), False, False
    Dcombos.GetAccountingCodes Me.Account_code3(1), False, False
    Dcombos.GetAccountingCodes Me.Account_code4(1), False, False
    
    Dcombos.GetSalesRepData Me.SalesPersonid
    Dcombos.GetSalesRepDatapurchase Me.PurchasePersonid
    'Dcombos.GetSalesRepDatapurchase Me.PurchasePersonid

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
Cmd(5).Caption = "Search"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
  Cmd(7).Caption = "Prient All"
  Cmd(8).Caption = "Prient"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
Lbl(17).Caption = "Sales Peson"
Lbl(18).Caption = "Purcahse Peson"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
Frame10.Caption = "Users Data"
CmdRemove.Caption = "Remove Line"
    Me.Caption = "Stores Data"
     Label1(0).Caption = Me.Caption
    Lbl(0).Caption = "Code"
    Lbl(1).Caption = "Name AR"
    Lbl(13).Caption = "Name EN"

    Lbl(6).Caption = "Address"
    Lbl(5).Caption = "Tel"
    Lbl(7).Caption = "Manger"
    Lbl(8).Caption = "Stock  Acc."
    Lbl(9).Caption = "loss and damage Acc."
    Lbl(10).Caption = "Inventory adjustments Acc."
    Label3.Caption = "Branch Data"

    Lbl(11).Caption = "Gifts & Samples Acc."

    Lbl(3).Caption = "Remarks"
    Lbl(2).Caption = "Current rec."
    Lbl(4).Caption = "Rec. Count."

    C1Tab1.TabCaption(0) = "Stores Data"
    C1Tab1.TabCaption(1) = "Locations"
    C1Tab1.TabCaption(2) = "Internal Rule"
    C1Tab1.TabCaption(3) = "Users Data"
    
    Lbl(14).Caption = C1Tab1.TabCaption(2)
    Lbl(15).Caption = "Location Name"
    Cmd(20).Caption = "ADD"
    Cmd(21).Caption = "Remove"

    With Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Location")) = "Locations"
    End With
 
 
   With VSFlexGrid1
        .TextMatrix(0, .ColIndex("LineNo")) = "Ser"
       .TextMatrix(0, .ColIndex("fullcode")) = "User Code"
        .TextMatrix(0, .ColIndex("name")) = "User Name"
    End With
 Label2(0).Caption = "NO"
 Lbl(0).Caption = "Code0"
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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



Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = "»Ì«‰«  «·„Œ«“‰"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.XPTxtStoreID.locked = True
            Me.XPTxtStoreName.locked = True
            Me.XPTxtStoreAddress.locked = True
            Me.XPTxtStorePhone.locked = True
            Me.XPMTxtRemark.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

        Case "N"
            '        Me.Caption = "»Ì«‰«  «·„Œ«“‰( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtStoreID.locked = True
            Me.XPTxtStoreName.locked = False
            Me.XPTxtStoreAddress.locked = False
            Me.XPTxtStorePhone.locked = False
            Me.XPMTxtRemark.locked = False

        Case "E"
            '        Me.Caption = "»Ì«‰«  «·„Œ«“‰(  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            Me.XPTxtStoreID.locked = True
            Me.XPTxtStoreName.locked = False
            Me.XPTxtStoreAddress.locked = False
            Me.XPTxtStorePhone.locked = False
            Me.XPMTxtRemark.locked = False
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
  On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.Find "StoreID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
    Account_code1(1).Enabled = False
    Account_code2(1).Enabled = False
    Account_code3(1).Enabled = False
    Account_code4(1).Enabled = False
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), 0, val(rs("BranchId").value))
    XPTxtStoreID.text = IIf(IsNull(rs("StoreID").value), "", val(rs("StoreID").value))
    Me.TXTCode.text = IIf(IsNull(rs("Code").value), "", (rs("Code").value))

    XPTxtStoreName.text = IIf(IsNull(rs("StoreName").value), "", Trim(rs("StoreName").value))
    XPTxtStoreNamee.text = IIf(IsNull(rs("StoreNamee").value), "", Trim(rs("StoreNamee").value))
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    XPTxtStoreAddress.text = IIf(IsNull(rs("StoreAdress").value), "", Trim(rs("StoreAdress").value))
    XPTxtStorePhone.text = IIf(IsNull(rs("StorePhone").value), "", Trim(rs("StorePhone").value))
    XPMTxtRemark.text = IIf(IsNull(rs("Remarks").value), "", Trim(rs("Remarks").value))
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    Me.SalesPersonid.BoundText = IIf(IsNull(rs("SalesPersonid").value), "", rs("SalesPersonid").value)
    Me.PurchasePersonid.BoundText = IIf(IsNull(rs("PurchasePersonid").value), "", rs("PurchasePersonid").value)
    
    Me.Account_code1(0).BoundText = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
    
    
    Me.Account_code2(0).BoundText = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
    
    
    Me.Account_code1(1).BoundText = IIf(IsNull(rs("Account_Code0").value), "", rs("Account_Code0").value)
    If Account_code1(1).text = "" Then
        Account_code1(1).BoundText = get_account_code_branch(0, dcBranch.BoundText)
    End If
    
    Me.Account_code3(1).BoundText = IIf(IsNull(rs("Account_Code22").value), "", rs("Account_Code22").value)
    If Account_code3(1).text = "" Then
        Account_code3(1).BoundText = get_account_code_branch(11, dcBranch.BoundText)
    End If
        
        Me.Account_code2(1).BoundText = IIf(IsNull(rs("Account_Code11").value), "", rs("Account_Code11").value)
    If Account_code2(1).text = "" Then
        Dim currentidAccount As Integer
        If SystemOptions.eachStoreHaveLossAccount = True Then
              currentidAccount = 10
        Else
              currentidAccount = 75
        End If
        Account_code2(1).BoundText = get_account_code_branch(currentidAccount, dcBranch.BoundText)
    End If
    Me.Account_code4(1).BoundText = IIf(IsNull(rs("Account_Code33").value), "", rs("Account_Code33").value)
    If Account_code4(1).text = "" Then
      If SystemOptions.eachStoreHaveGiftAccount = True Then
            currentidAccount = 17
      Else
            currentidAccount = 76
      End If
    
        Account_code4(1).BoundText = get_account_code_branch(currentidAccount, dcBranch.BoundText)
   End If


   
   
    
   
    Me.Account_code3(0).BoundText = IIf(IsNull(rs("Account_Code2").value), "", rs("Account_Code2").value)
    
    
    Me.Account_code4(0).BoundText = IIf(IsNull(rs("Account_Code3").value), "", rs("Account_Code3").value)
    

    If rs("linked").value = True Then
        Me.Chk1.value = vbChecked
    Else
        Me.Chk1.value = vbUnchecked
    End If

    

    If rs("IsLab").value = True Then
        Me.chkIsLab.value = vbChecked
    Else
        Me.chkIsLab.value = vbUnchecked
    End If

        If rs("IsNotCreateEntry").value = True Then
        Me.chkIsNotCreateEntry.value = vbChecked
    Else
        Me.chkIsNotCreateEntry.value = vbUnchecked
    End If


    '»Ì«‰«  «·„” Œœ„Ì‰ ›Ì «·„Œ“‰
    Dim RsEmployee As ADODB.Recordset
    Set RsEmployee = New ADODB.Recordset
     Dim StrSQL As String
 Dim i As Integer
'      StrSQL = "SELECT     TOP 100 PERCENT dbo.TblUsersStores.userid, dbo.TblUsers.UserName"
'StrSQL = StrSQL + "  FROM         dbo.TblUsersStores INNER JOIN"
'StrSQL = StrSQL + "  dbo.TblUsers ON dbo.TblUsersStores.userid = dbo.TblUsers.UserID"
'StrSQL = StrSQL + "  Where (dbo.TblUsersStores.StoreID = " & val(Me.XPTxtStoreID.text) & ")"
'StrSQL = StrSQL + "  ORDER BY dbo.TblUsersStores.id"


StrSQL = " SELECT     TOP 100 PERCENT dbo.TblUsersStores.userid, dbo.TblUsers.UserName, dbo.TblEmployee.Fullcode"
StrSQL = StrSQL + "  FROM         dbo.TblUsersStores INNER JOIN"
StrSQL = StrSQL + "  dbo.TblUsers ON dbo.TblUsersStores.userid = dbo.TblUsers.UserID LEFT OUTER JOIN"
StrSQL = StrSQL + "  dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL + "  Where (dbo.TblUsersStores.StoreID = " & val(Me.XPTxtStoreID.text) & ")"
StrSQL = StrSQL + "  ORDER BY dbo.TblUsersStores.id"


    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then

        With Me.VSFlexGrid1
            .rows = .FixedRows + RsEmployee.RecordCount

            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsEmployee("userid").value), 0, val(RsEmployee("userid").value))
  .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(RsEmployee("fullcode").value), 0, (RsEmployee("fullcode").value))
  

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsEmployee("UserName").value), "", RsEmployee("UserName").value)
                Else
                    .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(RsEmployee("UserName").value), "", RsEmployee("UserName").value)
                End If

                RsEmployee.MoveNext
            Next i

            ' .AutoSize 0, .Cols - 1, False
                    
        End With
Else
VSFlexGrid1.rows = 1
    End If

    
        
       
        

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub


Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then
        Exit Sub
    End If
     
    If VSFlexGrid1.rows > 1 Then
        If VSFlexGrid1.rows = 2 Then
            Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.VSFlexGrid1.rows > 1 Then
                If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                    Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                End If
            End If
        End If
    End If
            
    ReLineGrid

End Sub

 

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid1

        Select Case .ColKey(Col)
 
            Case "name"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
             
          
  StrSQL = "SELECT     TOP 100 PERCENT dbo.TblUsers.UserName, dbo.TblEmployee.Fullcode, dbo.TblUsers.UserID"
StrSQL = StrSQL & "  FROM         dbo.TblUsersStores RIGHT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblUsers ON dbo.TblUsersStores.userid = dbo.TblUsers.UserID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.TblUsers.Empid = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "  Where (dbo.TblUsers.UserID = " & val(StrAccountCode) & ")"
StrSQL = StrSQL & "  ORDER BY dbo.TblUsersStores.id"
                    
                    Set rs = Nothing
                
                    If StrAccountCode <> "" Then
                            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                            If Not (rs.BOF Or rs.EOF) Then
                                     .TextMatrix(Row, .ColIndex("fullcode")) = _
                                    IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)
                             
                            End If
                    End If
            
                 
         
        End Select
   
        If Row = .rows - 1 Then
    
            .rows = .rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "name"
                Exit Sub
        End Select

    End With

    VSFlexGrid1.ComboList = ""

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "name"
                StrSQL = "select * from TblUsers"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "UserName", "UserID")
                Else
                    StrComboList = VSFlexGrid1.BuildComboList(rs, "UserName", "UserID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
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

Private Sub SaveData()
    'On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean

    If Me.TxtModFlg.text <> "R" Then
        If Trim(dcBranch.BoundText) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify Departement"
            Else
                Msg = "ÌÃ»  ÕœÌœ «”„    «·›—⁄"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            dcBranch.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If XPTxtStoreName.text = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "Specify Store Name ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·„Œ“‰ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
        
            XPTxtStoreName.SetFocus
            Exit Sub
        End If

        Select Case Me.TxtModFlg.text

            Case "N"
                StrSQL = "select * From  TblStore where StoreName='" & Trim(XPTxtStoreName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    Msg = "ÌÊÃœ „Œ“‰ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Œ“‰"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTxtStoreName.SetFocus
                    Exit Sub
                End If

            Case "E"
                StrSQL = "select * From  TblStore where StoreName='" & Trim(XPTxtStoreName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("StoreID").value <> val(XPTxtStoreID.text) Then
                        Msg = "ÌÊÃœ „Œ“‰ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Œ“‰"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtStoreName.SetFocus
                        Exit Sub
                    End If
                End If

        End Select
    
        Dim Account_Code_dynamic As String
        Dim Account_Code_dynamic1 As String
        Dim Account_Code_dynamic2 As String
        Dim Account_Code_dynamic3 As String
         
        Account_Code_dynamic = get_account_code_branch(0, my_branch)
        Account_Code_dynamic = Account_code1(1).BoundText
        If Account_Code_dynamic = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created", vbCritical
            End If

            GoTo ErrTrap
        Else

            If Account_Code_dynamic = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "Inventory Account not defined on this branch", vbCritical
                End If
        
                GoTo ErrTrap
         
            End If
        End If






    If SystemOptions.StoreAccountHaveSettelment = False Then

        Account_Code_dynamic2 = get_account_code_branch(11, my_branch)
        Account_Code_dynamic2 = Account_code2(1).BoundText
        If Account_Code_dynamic2 = "NO branch" Then
        
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created", vbCritical
            End If
        
            GoTo ErrTrap
        Else

            If Account_Code_dynamic2 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» «· ”ÊÌ«  «·Ã—œÌ… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "Inventory Adjustment Account not defined on this branch", vbCritical
                End If

                GoTo ErrTrap
         
            End If
        End If
        
        
 End If

Dim currentidAccount As Integer
      If SystemOptions.eachStoreHaveLossAccount = True Then
            currentidAccount = 10
      Else
      currentidAccount = 75
        End If


        Account_Code_dynamic1 = get_account_code_branch(currentidAccount, my_branch)
        Account_Code_dynamic1 = Account_code2(1).BoundText
        If Account_Code_dynamic1 = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created", vbCritical
            End If
        
            GoTo ErrTrap
        Else

            If Account_Code_dynamic1 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ›ﬁœ Ê ·›  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "damage and loass Account not defined on this branch", vbCritical
                End If

                GoTo ErrTrap
         
            End If
        End If
        

      If SystemOptions.eachStoreHaveGiftAccount = True Then
            currentidAccount = 17
      Else
      currentidAccount = 76
        End If


        
        Account_Code_dynamic3 = get_account_code_branch(currentidAccount, my_branch)
        Account_Code_dynamic3 = Account_code4(1).BoundText
        If Account_Code_dynamic3 = "NO branch" Then
          
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created", vbCritical
            End If
        
            GoTo ErrTrap
        Else

            If Account_Code_dynamic3 = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» Âœ«Ì« Ê⁄Ì‰«    ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "Gifts and Sample Adjustment Account not defined on this branch", vbCritical
                End If

                GoTo ErrTrap
         
            End If
        End If
        
        Dim last_account As Boolean
        Dim link_account As Boolean
        Dim rsOut As New ADODB.Recordset
        Set rsOut = New ADODB.Recordset
        rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

        If Not (rsOut.EOF Or rsOut.BOF) Then
 
            If rsOut!opt_group = True Then
                If rsOut!Opt_Inventory_create_account = 1 Then '—»ÿ „Œ«“‰ Ê›—€
                    last_account = True
                    link_account = False
                ElseIf rsOut!opt_inv_and_branch_create_account = 1 Then
                    last_account = False
                    link_account = True
                End If
     
            Else
                last_account = True
                Me.Chk1.value = Unchecked
            End If
        End If
    
        Select Case Me.TxtModFlg.text

            Case "N"
                rs.AddNew
                rs("StoreID").value = val(XPTxtStoreID.text)
        
                If detect_inventory_work_type = 2 Or detect_inventory_work_type = 3 Then
       
       
       
   '                 rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, "  «·„Œ“Ê‰ «·”·⁄Ì   " & XPTxtStoreName.text, last_account, False, XPTxtStoreNamee.text & "   -Inventory")
   '                 rs("Account_Code1").value = ModAccounts.AddNewAccount(Account_Code_dynamic1, "  ›ﬁœ Ê ·›   " & XPTxtStoreName.text, last_account, False, XPTxtStoreNamee.text & " - Loss and damage")
   '                 rs("Account_Code2").value = ModAccounts.AddNewAccount(Account_Code_dynamic2, "   «· ”ÊÌ«  «·Ã—œÌ…   " & XPTxtStoreName.text, last_account, False, XPTxtStoreNamee.text & " - Inventory adjustments")
   '                 rs("Account_Code3").value = ModAccounts.AddNewAccount(Account_Code_dynamic3, "  Âœ«Ì« Ê⁄Ì‰«      " & XPTxtStoreName.text, last_account, False, XPTxtStoreNamee.text & "  -Gifts & Samples")
        
        
        
 Dim X As String
  '»Ì«‰«  ·Õ”«»« 
          If SystemOptions.StoreAccountHaveSettelment = True Then
        X = ModAccounts.AddNewAccount(Account_Code_dynamic, XPTxtStoreName.text, False, False, XPTxtStoreNamee.text & "   -Inventory")
        rs("ParetnAccount").value = X
        rs("Account_Code").value = ModAccounts.AddNewAccount(X, " «·„Œ“Ê‰ «·”·⁄Ì " & XPTxtStoreName.text, True, False, XPTxtStoreNamee.text & " - Inventory adjustments")
        rs("Account_Code2").value = ModAccounts.AddNewAccount(X, "    «· ”ÊÌ«  «·Ã—œÌ…  " & XPTxtStoreName.text, True, False, XPTxtStoreNamee.text & " Accumulated depreciation")
    Else
         rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, " «·„Œ“Ê‰ «·”·⁄Ì " & XPTxtStoreName.text, True, False, XPTxtStoreNamee.text & " - Inventory Adjustments")
        rs("Account_Code2").value = ModAccounts.AddNewAccount(Account_Code_dynamic2, "    «· ”ÊÌ«  «·Ã—œÌ…  " & XPTxtStoreName.text, True, False, XPTxtStoreNamee.text & " Accumulated Depreciation")
        
     End If
                 

                  
       If Me.chkIsLab.value = vbChecked Then
            rs("IsLab").value = 1

        ElseIf Me.chkIsLab.value = vbUnchecked Then
            rs("IsLab").value = 0
        End If
        
        
       If Me.chkIsNotCreateEntry.value = vbChecked Then
            rs("IsNotCreateEntry").value = 1

        ElseIf Me.chkIsNotCreateEntry.value = vbUnchecked Then
            rs("IsNotCreateEntry").value = 0
        End If
        
        
              
    If SystemOptions.eachStoreHaveLossAccount = True Then
        rs("Account_Code1").value = ModAccounts.AddNewAccount(Account_Code_dynamic1, "  ›ﬁœ Ê ·›   " & XPTxtStoreName.text, last_account, False, XPTxtStoreNamee.text & " - Loss and damage")
    Else
    rs("Account_Code1").value = Account_Code_dynamic1
    End If
    
    
        If SystemOptions.eachStoreHaveGiftAccount = True Then
        rs("Account_Code3").value = ModAccounts.AddNewAccount(Account_Code_dynamic3, "  Âœ«Ì« Ê⁄Ì‰«      " & XPTxtStoreName.text, last_account, False, XPTxtStoreNamee.text & "  -Gifts & Samples")
    Else
    rs("Account_Code3").value = Account_Code_dynamic3
    End If
    
    
    
        
                End If
       
        End Select

        Cn.BeginTrans
        BeginTrans = True
    
        rs("Code").value = Trim(TXTCode.text)
        rs("StoreName").value = Trim(XPTxtStoreName.text)
        rs("StoreNamee").value = Trim(XPTxtStoreNamee.text)
        
        rs("StoreAdress").value = IIf(XPTxtStoreAddress.text = "", "", Trim(XPTxtStoreAddress.text))
        rs("StorePhone").value = IIf(XPTxtStorePhone.text = "", "", Trim(XPTxtStorePhone.text))
        rs("Remarks").value = IIf(XPMTxtRemark.text = "", "", Trim(XPMTxtRemark.text))
        rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", "", val(DcboEmp.BoundText))
        rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
        
        rs("SalesPersonid").value = IIf(SalesPersonid.BoundText = "", 0, val(SalesPersonid.BoundText))
        rs("PurchasePersonid").value = IIf(PurchasePersonid.BoundText = "", 0, val(PurchasePersonid.BoundText))
        
        
        rs("Account_Code0").value = IIf(Account_code1(1) = "", "", Trim(Account_code1(1).BoundText))
        rs("Account_Code22").value = IIf(Account_code3(1).BoundText = "", "", Trim(Account_code3(1).BoundText))
        rs("Account_Code11").value = IIf(Account_code2(1).BoundText = "", "", Trim(Account_code2(1).BoundText))
        rs("Account_Code33").value = IIf(Account_code4(1).BoundText = "", "", Trim(Account_code4(1).BoundText))
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
        rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", "", val(DcboEmp.BoundText))
        
        rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
                  
       If Me.chkIsLab.value = vbChecked Then
            rs("IsLab").value = 1

        ElseIf Me.chkIsLab.value = vbUnchecked Then
            rs("IsLab").value = 0
        End If
        
                          
       If Me.chkIsNotCreateEntry.value = vbChecked Then
            rs("IsNotCreateEntry").value = 1

        ElseIf Me.chkIsNotCreateEntry.value = vbUnchecked Then
            rs("IsNotCreateEntry").value = 0
        End If
        
       
        
        If Me.Chk1.value = vbChecked Then
            rs("linked").value = 1

        ElseIf Me.Chk1.value = vbUnchecked Then
            rs("linked").value = 0
        End If

        If Me.TxtModFlg.text = "E" Then
            
            
                 If Not IsNull(rs("ParetnAccount").value) Then
            ModAccounts.EditAccount rs("ParetnAccount").value, Me.XPTxtStoreName.text, Trim(XPTxtStoreNamee.text), , , , , , , , , , , , , , , , , False
        End If
            
 
        
            
            
            If Not IsNull(rs("Account_Code").value) Then
                ModAccounts.EditAccount rs("Account_Code").value, Me.XPTxtStoreName.text & "  «·„Œ“Ê‰ «·”·⁄Ì   ", XPTxtStoreNamee.text & "   -Inventory", , , , , , , , , , , , , , , , , last_account
            End If
                  
                  
               If Not IsNull(rs("Account_Code2").value) Then
                ModAccounts.EditAccount rs("Account_Code2").value, Me.XPTxtStoreName.text & "   «· ”ÊÌ«  «·Ã—œÌ…   ", XPTxtStoreNamee.text & " - Inventory adjustments", , , , , , , , , , , , , , , , , last_account
            End If
            
            
            If SystemOptions.eachStoreHaveLossAccount = True Then
                    If Not IsNull(rs("Account_Code1").value) Then
                        ModAccounts.EditAccount rs("Account_Code1").value, Me.XPTxtStoreName.text & "  ›ﬁœ Ê ·›   ", XPTxtStoreNamee.text & " - Loss and damage", , , , , , , , , , , , , , , , , last_account
                    End If
                    
            End If
            
                  
      
                  If SystemOptions.eachStoreHaveGiftAccount = True Then
            If Not IsNull(rs("Account_Code3").value) Then
                ModAccounts.EditAccount rs("Account_Code3").value, Me.XPTxtStoreName.text & "  Âœ«Ì« Ê⁄Ì‰«      ", XPTxtStoreNamee.text & "  -Gifts & Samples", , , , , , , , , , , , , , , , , last_account
            End If
            
        End If
        
        
        
        
        End If
     

        
        

        
        
        rs.update
        
        
       'update tblusers
             Dim RsEmployee As ADODB.Recordset
        Dim i As Integer
 
        If Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From TblUsersStores Where storeId=" & val(Me.XPTxtStoreID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If

        If Me.VSFlexGrid1.rows <> 1 Then
            Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersStores", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
            
            If VSFlexGrid1.rows > 2 Then
                VSFlexGrid1.rows = VSFlexGrid1.rows - 1
            End If

            For i = 1 To Me.VSFlexGrid1.rows - 1
                RsEmployee.AddNew
                RsEmployee("storeId").value = val(Me.XPTxtStoreID.text)
                RsEmployee("userid").value = val(VSFlexGrid1.TextMatrix(i, VSFlexGrid1.ColIndex("id")))
                  
                RsEmployee.update
            Next i

        End If


        
        Cn.CommitTrans
        
                Dcombos.GetAccountingCodes Me.Account_code1(0)
        Dcombos.GetAccountingCodes Me.Account_code2(0)
        Dcombos.GetAccountingCodes Me.Account_code3(0)
        Dcombos.GetAccountingCodes Me.Account_code4(0)


    Dcombos.GetAccountingCodes Me.Account_code1(1), False, False
    Dcombos.GetAccountingCodes Me.Account_code2(1), False, False
    Dcombos.GetAccountingCodes Me.Account_code3(1), False, False
    Dcombos.GetAccountingCodes Me.Account_code4(1), False, False
        
        Me.Account_code1(0).BoundText = rs("Account_Code").value
        Me.Account_code2(0).BoundText = rs("Account_Code1").value
        Me.Account_code3(0).BoundText = rs("Account_Code2").value
        Me.Account_code4(0).BoundText = rs("Account_Code3").value
        
        
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    Account_code1(1).Enabled = False
    Account_code2(1).Enabled = False
    Account_code3(1).Enabled = False
    Account_code4(1).Enabled = False
        If link_account = True Then
            If create_accounts(XPTxtStoreID.text) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ —»ÿ «·„Œ“‰ „⁄ «·„Ã„Ê⁄« "
                Else
                    MsgBox "joined with group"
                End If

            Else

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ —»ÿ «·„Œ“‰ „⁄ «·„Ã„Ê⁄« "
                Else
                    MsgBox "Not joined with group"
                End If
            End If
                
        End If

        CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„Œ“‰" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Data was Save , do you want to add another Y/n" & CHR(13)
        
                End If
            
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox "Changes was Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Function create_accounts(inv_id As Integer) As Boolean
    Dim rsOut As New ADODB.Recordset
    Dim Current_case As Integer
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!opt_group = False Then
            create_accounts = False
            Exit Function
        End If

    Else
        create_accounts = False
        Exit Function
    End If

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    sql = "Select * from Groups where not(ParentID is null)"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
        Exit Function
    End If

    For i = 1 To Rs3.RecordCount

        If create_inventory_group(inv_id, Rs3("GroupID").value, Rs3("GroupName").value) = True Then
        End If

        Rs3.MoveNext
    Next i

    Rs3.Close
    create_accounts = True
End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "StoreID='" & val(XPTxtStoreID.text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim StrAccountCode2 As String
    Dim StrAccountCode3 As String
    Dim RsTemp1 As New ADODB.Recordset
Dim ParetnAccount As String

    On Error GoTo ErrTrap

    If XPTxtStoreID.text <> "" Then
        StrSQL = "select * From Transactions where StoreID=" & XPTxtStoreID.text
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            Msg = "·« Ì„ﬂ‰ Õ–›  »Ì«‰«  Â–« «·„Œ“‰" & CHR(13)
            Msg = Msg + " „  ”ÃÌ· »⁄÷ «·⁄„·Ì«  ⁄·Ï Â–« «·„Œ“‰"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
        
        

         StrAccountCode = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value) '„Œ“Ê‰
        
       StrAccountCode2 = IIf(IsNull(rs("Account_Code2").value), "", rs("Account_Code2").value) ' ”ÊÌ« 
       
        
       '›ﬁœ
        
        
            If SystemOptions.StoreAccountHaveSettelment = True Then
        
         ParetnAccount = IIf(IsNull(rs("ParetnAccount").value), "", rs("ParetnAccount").value) '«·»
        Else
        ParetnAccount = "a000"
        End If
        

        If SystemOptions.eachStoreHaveLossAccount = True Then
        
         StrAccountCode1 = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value) '›ﬁœ
        Else
        StrAccountCode1 = "a000"
        End If
        
        If SystemOptions.eachStoreHaveGiftAccount = True Then
        StrAccountCode3 = IIf(IsNull(rs("Account_Code3").value), "", rs("Account_Code3").value) 'Âœ«Ì«
        Else
        StrAccountCode3 = "a000"
        End If
        StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where Account_Code='" & StrAccountCode & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode1 & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode2 & "'"
        StrSQL = StrSQL & " Or  Account_Code='" & StrAccountCode3 & "'"
        
        RsTemp1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp1.EOF Or RsTemp1.BOF) Then
            Msg = "·« Ì„ﬂ‰ Õ–›  »Ì«‰«  Â–« «·„Œ“‰" & CHR(13)
            Msg = Msg + " „  ”ÃÌ· »⁄÷ «·ﬁÌÊœ ⁄·Ï Â–« «·„Œ“‰"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    
        Msg = "”Ì „ Õ–› »Ì«‰«  «·„Œ“‰ —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtStoreID.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                '    CuurentLogdata ("D")
            
                '        rs.delete
    
                If ModAccounts.DeleteAccount(ParetnAccount, True) = True Or ModAccounts.DeleteAccount(StrAccountCode, True) = True And ModAccounts.DeleteAccount(StrAccountCode1, True) = True And ModAccounts.DeleteAccount(StrAccountCode2, True) = True And ModAccounts.DeleteAccount(StrAccountCode3, True) = True Then
                    CuurentLogdata ("D")
                    rs.delete
             
                    Msg = " „  ⁄„·Ì… «·Õ–›."
                    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
                Else
                    GoTo ErrTrap
                End If
            
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
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    On Error GoTo ErrTrap
    Dim Wrap As String
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  „Œ“‰ ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·„Œ“‰" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·„Œ“‰ «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·„Œ“‰" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ „Œ“‰" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·„Œ«“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
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
        If Cmd(0).Enabled = False Then
            Exit Sub
        End If
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then
            Exit Sub
        End If
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then
            Exit Sub
        End If
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then
            Exit Sub
        End If
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then
            Exit Sub
        End If
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (6)
        End If
    End If

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

Function CuurentLogdata(Optional Currentmode As String)

 ScreenNameArabic = "  »Ì«‰«  «·„Œ«“‰ "
    ScreenNameEnglish = " Store Data  "


    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " ﬂÊœ «·„Œ“‰    " & TXTCode.text & CHR(13) & " «”„ «·„Œ“‰ " & XPTxtStoreName & CHR(13) & " «·›—⁄   " & dcBranch.text & CHR(13) & "«„Ì‰ «·„” Êœ⁄  " & XPTxtStoreName & CHR(13) & " «· ·Ì›Ê‰  " & XPTxtStorePhone.text & CHR(13) & " Õ”«» «·„Œ“Ê‰  " & Account_code1(0) & CHR(13) & " Õ”«» Âœ«Ì« Ê⁄Ì‰«   " & Account_code2(0).text & CHR(13) & " Õ”«» ›ﬁœ Ê ·›   " & Account_code3(0).text & CHR(13) & " Õ”«» «· ”ÊÌ«  «·Ã—œÌ…  " & Account_code4(0).text
                      
    '       If Chk1.value = vbChecked Then
    '        LogTextA = LogTextA & Chr(13) & "  „ «·—»ÿ «·„Œ“‰ »«·„Ã„Ê⁄«   "
    '       End If
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Store Name" & XPTxtStoreNamee & CHR(13) & " Branch   " & dcBranch.text & CHR(13) & " Employee " & XPTxtStoreName & CHR(13) & " Tel  " & XPTxtStorePhone.text & CHR(13) & " Inventory Acc. " & Account_code1(0) & CHR(13) & " Gifts and Sample Acc.  " & Account_code2(0).text & CHR(13) & "  Lose And Dep. Acc   " & Account_code3(0).text & CHR(13) & "  Adjust. Acc  " & Account_code4(0).text
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function



Private Sub XPTxtStoreName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub



Private Sub XPTxtStoreNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub



