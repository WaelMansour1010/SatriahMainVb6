VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmExpensesSearch 
   Appearance      =   0  'Flat
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " «·Œ“‰   Ê «·⁄Âœ"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10275
   Icon            =   "ExpensesSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   10275
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰› «·„—«œ «·»ÕÀ ⁄‰Â ÌÕ ÊÏ ⁄·Ï Â–« «·’‰› ﬂ«Õœ „·Õﬁ« Â"
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
      Height          =   885
      Left            =   4740
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   6810
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰› «·„—«œ «·»ÕÀ ⁄‰Â ÌÕ ÊÏ ⁄·Ï Â–« «·’‰› ﬂ«Õœ „ﬂÊ‰« Â"
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
      Height          =   885
      Left            =   4110
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   7860
      Width           =   6495
   End
   Begin VB.ComboBox CboArchive 
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6300
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox CboGuar 
      Height          =   315
      Left            =   2160
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5850
      Width           =   1305
   End
   Begin VB.ComboBox CboAttachedItem 
      Height          =   315
      Left            =   4470
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   6180
      Width           =   1095
   End
   Begin VB.ComboBox CboAssbliedItem 
      Height          =   315
      Left            =   2160
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5700
      Width           =   1305
   End
   Begin VB.ComboBox CboItemType 
      Height          =   315
      Left            =   4380
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5970
      Width           =   1215
   End
   Begin VB.ComboBox CboSerial 
      Height          =   315
      ItemData        =   "ExpensesSearch.frx":030A
      Left            =   30
      List            =   "ExpensesSearch.frx":030C
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5580
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox XPTxtItemCode 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2610
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   6615
      Width           =   1395
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1950
      TabIndex        =   7
      Top             =   4905
      Width           =   945
      _ExtentX        =   1667
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
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   930
      TabIndex        =   8
      Top             =   4905
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   30
      TabIndex        =   9
      Top             =   4905
      Width           =   855
      _ExtentX        =   1508
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
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   4875
      Left            =   60
      TabIndex        =   20
      Top             =   -30
      Width           =   10170
      _cx             =   17939
      _cy             =   8599
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
      BackColor       =   12648447
      ForeColor       =   -2147483630
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "«·„’—Ê›« |«·»‰Êﬂ| «·Œ“‰   Ê «·⁄Âœ|«·≈Ì—«œ« "
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
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   4500
         Index           =   0
         Left            =   45
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   45
         Width           =   10080
         _cx             =   17780
         _cy             =   7938
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
         Begin VB.TextBox TxtItemName 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   0
            Left            =   6345
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   3240
            Width           =   2535
         End
         Begin VB.TextBox TxtItemID 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   0
            Left            =   6270
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   2850
            Width           =   2610
         End
         Begin VB.ComboBox CboNameSearch 
            Height          =   315
            Index           =   0
            Left            =   0
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   3240
            Width           =   1575
         End
         Begin VB.ComboBox CboItemCodeSearch 
            Height          =   315
            Index           =   0
            Left            =   0
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   2850
            Width           =   1575
         End
         Begin VB.TextBox tXTAccount_Serial 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   0
            Left            =   2655
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   2820
            Width           =   2655
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   4440
            Index           =   0
            Left            =   14070
            TabIndex        =   22
            Top             =   345
            Width           =   9900
            _cx             =   17462
            _cy             =   7832
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
            Rows            =   50
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ExpensesSearch.frx":030E
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
         Begin VSFlex8UCtl.VSFlexGrid Fg 
            Height          =   2745
            Index           =   0
            Left            =   0
            TabIndex        =   30
            Top             =   30
            Width           =   10140
            _cx             =   17886
            _cy             =   4842
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ExpensesSearch.frx":03CE
            ScrollTrack     =   -1  'True
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
         Begin MSDataListLib.DataCombo DCboGroupName 
            Height          =   315
            Index           =   0
            Left            =   2595
            TabIndex        =   31
            Top             =   3240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·›∆…"
            Height          =   285
            Index           =   3
            Left            =   5130
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   3270
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„’—Ê›"
            Height          =   345
            Index           =   1
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   3240
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„’—Ê›"
            Height          =   345
            Index           =   4
            Left            =   8940
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   2850
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã«· «·»ÕÀ"
            Height          =   345
            Index           =   8
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   3240
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã«· «·»ÕÀ"
            Height          =   345
            Index           =   11
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   2850
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·Õ”«»"
            Height          =   345
            Index           =   12
            Left            =   5250
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   2820
            Width           =   960
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   4500
         Index           =   1
         Left            =   10815
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   45
         Width           =   10080
         _cx             =   17780
         _cy             =   7938
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
         Begin VB.TextBox tXTAccount_SerialName 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   2640
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   3180
            Width           =   2655
         End
         Begin VB.TextBox tXTAccount_Serial 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   1
            Left            =   2655
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   2790
            Width           =   2655
         End
         Begin VB.ComboBox CboItemCodeSearch 
            Height          =   315
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   2820
            Width           =   1575
         End
         Begin VB.ComboBox CboNameSearch 
            Height          =   315
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   3210
            Width           =   1575
         End
         Begin VB.TextBox TxtItemID 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   1
            Left            =   6270
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   2820
            Width           =   2295
         End
         Begin VB.TextBox TxtItemName 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   1
            Left            =   6285
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   3660
            Width           =   2280
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   4440
            Index           =   1
            Left            =   14070
            TabIndex        =   24
            Top             =   345
            Width           =   9900
            _cx             =   17462
            _cy             =   7832
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
            Rows            =   50
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ExpensesSearch.frx":05E7
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
         Begin VSFlex8UCtl.VSFlexGrid Fg 
            Height          =   2745
            Index           =   1
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   10140
            _cx             =   17886
            _cy             =   4842
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ExpensesSearch.frx":06A7
            ScrollTrack     =   -1  'True
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
         Begin MSDataListLib.DataCombo DcCurrency 
            Height          =   315
            Left            =   4200
            TabIndex        =   75
            Top             =   3630
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Index           =   0
            Left            =   5700
            TabIndex        =   76
            Top             =   4110
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Õ”«»"
            Height          =   345
            Index           =   31
            Left            =   5205
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   3180
            Width           =   960
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   9285
            TabIndex        =   77
            Top             =   4110
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·Õ”«»"
            Height          =   345
            Index           =   18
            Left            =   5250
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   2790
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã«· «·»ÕÀ"
            Height          =   345
            Index           =   17
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   2820
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã«· «·»ÕÀ"
            Height          =   345
            Index           =   16
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   3210
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·»‰ﬂ"
            Height          =   345
            Index           =   15
            Left            =   8625
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   2820
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·»‰ﬂ"
            Height          =   345
            Index           =   14
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   3660
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„·…"
            Height          =   285
            Index           =   13
            Left            =   5130
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   3690
            Width           =   960
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   4500
         Index           =   2
         Left            =   11115
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   45
         Width           =   10080
         _cx             =   17780
         _cy             =   7938
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
         Begin VB.TextBox tXTAccount_Serial 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   2
            Left            =   2655
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   2790
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.ComboBox CboItemCodeSearch 
            Height          =   315
            Index           =   2
            Left            =   0
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   2820
            Width           =   1575
         End
         Begin VB.ComboBox CboNameSearch 
            Height          =   315
            Index           =   2
            Left            =   0
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   3210
            Width           =   1575
         End
         Begin VB.TextBox TxtItemID 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   2
            Left            =   6510
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   2820
            Width           =   1995
         End
         Begin VB.TextBox TxtItemName 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   2
            Left            =   6495
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   3210
            Width           =   2010
         End
         Begin VSFlex8UCtl.VSFlexGrid Fg 
            Height          =   2745
            Index           =   2
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   10140
            _cx             =   17886
            _cy             =   4842
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ExpensesSearch.frx":0905
            ScrollTrack     =   -1  'True
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
         Begin MSDataListLib.DataCombo DCboGroupName 
            Height          =   315
            Index           =   2
            Left            =   5835
            TabIndex        =   58
            Top             =   4050
            Visible         =   0   'False
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcEmp 
            Height          =   315
            Left            =   4410
            TabIndex        =   81
            Top             =   3660
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
            Index           =   1
            Left            =   2640
            TabIndex        =   83
            Top             =   3210
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   5445
            TabIndex        =   84
            Top             =   3210
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÊŸ›"
            Height          =   315
            Index           =   30
            Left            =   8850
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   3660
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·Õ”«»"
            Height          =   345
            Index           =   24
            Left            =   5250
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   2790
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã«· «·»ÕÀ"
            Height          =   345
            Index           =   23
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   2820
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã«· «·»ÕÀ"
            Height          =   345
            Index           =   22
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   3210
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·Œ“Ì‰… Ê«·⁄Âœ"
            Height          =   345
            Index           =   21
            Left            =   8505
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   2820
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Œ“Ì‰… Ê«·⁄Âœ"
            Height          =   345
            Index           =   20
            Left            =   8685
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   3210
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ”«» «·—∆Ì”Ì"
            Height          =   285
            Index           =   19
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   4050
            Visible         =   0   'False
            Width           =   1170
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   4500
         Index           =   3
         Left            =   11415
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   45
         Width           =   10080
         _cx             =   17780
         _cy             =   7938
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
         Begin VB.ComboBox CboItemCodeSearch 
            Height          =   315
            Index           =   3
            Left            =   0
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   2820
            Width           =   1575
         End
         Begin VB.ComboBox CboNameSearch 
            Height          =   315
            Index           =   3
            Left            =   0
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   3210
            Width           =   1575
         End
         Begin VB.TextBox TxtItemID 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   3
            Left            =   6270
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   2820
            Width           =   2610
         End
         Begin VB.TextBox TxtItemName 
            Alignment       =   1  'Right Justify
            Height          =   360
            Index           =   3
            Left            =   6285
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   3210
            Width           =   2565
         End
         Begin VSFlex8UCtl.VSFlexGrid Fg 
            Height          =   2745
            Index           =   3
            Left            =   0
            TabIndex        =   69
            Top             =   0
            Width           =   10140
            _cx             =   17886
            _cy             =   4842
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ExpensesSearch.frx":0B1F
            ScrollTrack     =   -1  'True
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
         Begin MSDataListLib.DataCombo DboParentAccount 
            Height          =   315
            Left            =   2700
            TabIndex        =   80
            Top             =   2820
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã«· «·»ÕÀ"
            Height          =   345
            Index           =   29
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   2820
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã«· «·»ÕÀ"
            Height          =   345
            Index           =   28
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   3210
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·«Ì—«œ"
            Height          =   345
            Index           =   27
            Left            =   8940
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   2820
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·«Ì—«œ"
            Height          =   345
            Index           =   26
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   3210
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ”«» «·—∆Ì”Ì"
            Height          =   285
            Index           =   25
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   2850
            Width           =   1320
         End
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·√—‘Ì›"
      Height          =   285
      Index           =   10
      Left            =   1410
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   5730
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·÷„«‰"
      Height          =   285
      Index           =   9
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   5850
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Ìﬁ«› «· ⁄«„·"
      Height          =   315
      Index           =   7
      Left            =   7770
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   6210
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " Ã„Ì⁄"
      Height          =   285
      Index           =   6
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   5700
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·’‰›"
      Height          =   285
      Index           =   5
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   5760
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Ìﬁ«› «· ⁄«„·"
      Height          =   315
      Index           =   2
      Left            =   1470
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5970
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblRes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5730
      Width           =   1905
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„—ﬂ“"
      Height          =   345
      Index           =   0
      Left            =   4020
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   5910
      Width           =   795
   End
End
Attribute VB_Name = "FrmExpensesSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo
Dim i As Long

Private m_RetrunType As Integer
Public Indx As Integer
Public Indx2 As Integer

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If
            Select Case Indx
            Case 0
                rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
            Case 1
                rs.Open Build_SqlBank, Cn, adOpenStatic, adLockReadOnly, adCmdText
            Case 2
                rs.Open Build_SqlBox, Cn, adOpenStatic, adLockReadOnly, adCmdText
            Case 3
                rs.Open Build_SqlRevenues, Cn, adOpenStatic, adLockReadOnly, adCmdText
            End Select

            If SystemOptions.UserInterface = ArabicInterface Then
                LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
    
            If rs.RecordCount < 1 Then
                Fg(Indx).Clear flexClearScrollable, flexClearEverything
                Fg(Indx).Rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.title
                End If

                Exit Sub
            End If

            Retrive
            Fg(Indx).SetFocus

        Case 1
            clear_all Me
            Fg(Indx).Clear flexClearScrollable, flexClearEverything

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub



Private Sub fg_Click(Index As Integer)
    On Error GoTo ErrTrap
 
    If Not Fg(Indx).TextMatrix(Fg(Indx).Row, 1) = "" Then
        If Me.RetrunType = 0 Then
            FrmExpensesType.Retrive val(Fg(Indx).TextMatrix(Fg(Indx).Row, 1))
        ElseIf Me.RetrunType = 20 Then
            FrmBanksData.Retrive val(Fg(Indx).TextMatrix(Fg(Indx).Row, 1))
        ElseIf Me.RetrunType = 21 Then
            FrmBoxesData.Retrive val(Fg(Indx).TextMatrix(Fg(Indx).Row, 1))
        ElseIf Me.RetrunType = 22 Then
            FrmRevenuesTypes.Retrive val(Fg(Indx).TextMatrix(Fg(Indx).Row, 1))
        ElseIf Me.RetrunType = 23 Then
            With frmserviceInvoice.Fg_Journal
                
                .TextMatrix(.Row, .ColIndex("AccountName")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 3)
                .TextMatrix(.Row, .ColIndex("AccountCode")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 5)
                .TextMatrix(.Row, .ColIndex("ExpensesID")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 1)
                .TextMatrix(.Row, .ColIndex("account_serial")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 6)
                frmserviceInvoice.Fg_Journal_StartEdit .Row, 4, False
            End With

            
        ElseIf Me.RetrunType = 1 Then

            With FrmExpenses5.Fg_Journal
                FrmExpenses5.Fg_Journal_StartEdit .Row, 4, False
                .TextMatrix(.Row, .ColIndex("AccountName")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 3)
                .TextMatrix(.Row, .ColIndex("AccountCode")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 5)
                .TextMatrix(.Row, .ColIndex("ExpensesID")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 1)
                .TextMatrix(.Row, .ColIndex("account_serial")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 6)
                
                .TextMatrix(.Row, .ColIndex("des")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 4)
                FrmExpenses5.Fg_Journal_AfterEdit .Row, .ColIndex("ExpensesID")
            End With
      ElseIf RetrunType = 986 Then

        

        For i = 0 To FrmEditUsers.ListBoxesAll.ListCount - 1
            If FrmEditUsers.ListBoxesAll.ItemData(i) = val(Fg(Indx).TextMatrix(Fg(Indx).Row, 1)) Then
                FrmEditUsers.ListBoxesAll.Selected(i) = True
                
                
            End If
        Next
       ElseIf Me.RetrunType = 1915 Then

            With RsExpenses.Fg_Journal
                RsExpenses.Fg_Journal_StartEdit .Row, 6, False
                .TextMatrix(.Row, .ColIndex("AccountName")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 3)
                .TextMatrix(.Row, .ColIndex("AccountCode")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 5)
                .TextMatrix(.Row, .ColIndex("ExpensesID")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 1)
                .TextMatrix(.Row, .ColIndex("des")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 4)
                RsExpenses.Fg_Journal_AfterEdit .Row, 6
            End With
            
            
        ElseIf Me.RetrunType = 2 Then
    
            With FrmExpenses3.Fg_Journal
                ' FrmExpenses3.Fg_Journal_StartEdit .Row, 4, False
                .TextMatrix(.Row, .ColIndex("AccountName")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 3)
                .TextMatrix(.Row, .ColIndex("AccountCode")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 5)
                .TextMatrix(.Row, .ColIndex("ExpensesID")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 1)
                .TextMatrix(.Row, .ColIndex("des")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 4)
                FrmExpenses3.Fg_Journal_AfterEdit .Row, 3
            End With
 
        ElseIf Me.RetrunType = 350 Then
    
            With FrmExpenses30.Fg_Journal
                ' FrmExpenses3.Fg_Journal_StartEdit .Row, 4, False
                .TextMatrix(.Row, .ColIndex("AccountName")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 3)
                .TextMatrix(.Row, .ColIndex("AccountCode")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 5)
                .TextMatrix(.Row, .ColIndex("ExpensesID")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 1)
                .TextMatrix(.Row, .ColIndex("des")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 4)
                .TextMatrix(.Row, .ColIndex("account_serial")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 6)
                FrmExpenses30.Fg_Journal_AfterEdit .Row, .ColIndex("ExpensesID")
            End With
 
        ElseIf Me.RetrunType = 3 Then

            With FrmProductionOrder.Fg_Journal
                ' FrmExpenses3.Fg_Journal_StartEdit .Row, 4, False
                .TextMatrix(.Row, .ColIndex("AccountName")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 3)
                .TextMatrix(.Row, .ColIndex("AccountCode")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 5)
                .TextMatrix(.Row, .ColIndex("ExpensesID")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 1)
                .TextMatrix(.Row, .ColIndex("des")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 4)
                FrmProductionOrder.Fg_Journal_AfterEdit .Row, 3
   
            End With
    
        ElseIf Me.RetrunType = 4 Then
    
            FrmExpenses2.DcCostCenter.BoundText = Fg(Indx).TextMatrix(Fg(Indx).Row, 2)
    
        ElseIf Me.RetrunType = 5 Then
            FrmPayments.DcCostCenter.BoundText = Fg(Indx).TextMatrix(Fg(Indx).Row, 2)
    
        ElseIf Me.RetrunType = 6 Then
    
            FrmCashing.DcCostCenter.BoundText = Fg(Indx).TextMatrix(Fg(Indx).Row, 2)
    
        ElseIf Me.RetrunType = 7 Then
    
            FrmEmployee.DcCostCenter.BoundText = Fg(Indx).TextMatrix(Fg(Indx).Row, 2)
    
        ElseIf Me.RetrunType = 8 Then
    
            FrmOpeningBalance.DCboItemsCode.BoundText = val(Fg(Indx).TextMatrix(Fg(Indx).Row, 1))
    
    
            ElseIf Me.RetrunType = 9 Then

            With FrmBillBuy.Fg_Journal
                ' FrmExpenses3.Fg_Journal_StartEdit .Row, 4, False
                .TextMatrix(.Row, .ColIndex("AccountName")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 3)
                .TextMatrix(.Row, .ColIndex("AccountCode")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 5)
                .TextMatrix(.Row, .ColIndex("ExpensesID")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 1)
                .TextMatrix(.Row, .ColIndex("des")) = Fg(Indx).TextMatrix(Fg(Indx).Row, 4)
                FrmBillBuy.Fg_Journal_AfterEdit .Row, 3
   
            End With
            
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    Fg(Indx).Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        Fg(Indx).Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With Fg(Indx)
            .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("Account_Serial")) = IIf(IsNull(rs("Account_Serial").value), "", (rs("Account_Serial").value))
                'Account_Serial
                .TextMatrix(Num, .ColIndex("ItemNum")) = IIf(IsNull(rs("id").value), "", val(rs("id").value))
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("KindNme")) = IIf(IsNull(rs("name").value), "", Trim(rs("name").value))
Else
                .TextMatrix(Num, .ColIndex("KindNme")) = IIf(IsNull(rs("namee").value), "", Trim(rs("namee").value))
End If
If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("Parent")) = IIf(IsNull(rs("parent_account").value), "", Trim(rs("parent_account").value))
Else
       .TextMatrix(Num, .ColIndex("Parent")) = ""
End If
.TextMatrix(Num, .ColIndex("account_code")) = IIf(IsNull(rs("account_code").value), "", Trim(rs("account_code").value))
            
            End With

            rs.MoveNext
        Next Num

        Fg(Indx).AutoSize 0, Fg(Indx).Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick(Index As Integer)
    fg_Click Index
    Unload Me
End Sub

Private Sub Form_Load()
5   On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    TabMain.TabEnabled(0) = False
    TabMain.TabEnabled(1) = False
    TabMain.TabEnabled(2) = False
    TabMain.TabEnabled(3) = False
    TabMain.TabVisible(0) = False
    TabMain.TabVisible(1) = False
    TabMain.TabVisible(2) = False
    TabMain.TabVisible(3) = False

    TabMain.TabEnabled(Indx) = True
    TabMain.TabVisible(Indx) = True
    
    TabMain.CurrTab = Indx
    
    Set Dcombos = New ClsDataCombos
    Dcombos.GetExpensesGroups Me.DCboGroupName(0)

    Set cSearchDcbo = New clsDCboSearch
    
    
    'cSearchDcbo.AllowWriting = False
  '  Set cSearchDcbo.Client = Me.DCboGroupName(Indx)
  
    Dim My_SQL As String
    'My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    'fill_combo dcBranch, My_SQL
    My_SQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, My_SQL
    
  '  Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch(0)
    Dcombos.GetBranches Me.Dcbranch(1)
  
    Dcombos.GetAccountingCodes Me.DboParentAccount, False, True, 2
    
      Dcombos.GetAccountingCodes Me.DCboGroupName(2), False, True
     Dcombos.GetEmployees Me.Dcemp
    Dim i As Integer
    If SystemOptions.UserInterface = ArabicInterface Then
        
        For i = 0 To CboItemCodeSearch.count - 1
            With Me.CboItemCodeSearch(i)
                .Clear
                .AddItem "»ÕÀ „ÿ«»ﬁ"
                .AddItem "»ÕÀ „‰ «·»œ«Ì…"
                .AddItem "»ÕÀ „‰ «·‰Â«Ì…"
                .AddItem "»ÕÀ ›Ï «Ï „ﬂ«‰"
            End With
            With Me.CboNameSearch(i)
                 .Clear
                 .AddItem "„‰ «Ê· «·√”„"
                 .AddItem "›Ï «Ï Ã“¡ „‰ «·√”„"
             End With
        Next
        With Me.CboSerial
            .Clear
            .AddItem "«·ﬂ·"
            .ItemData(0) = 0
            .AddItem "·Â ”Ì—Ì«·"
            .ItemData(1) = 1
            .AddItem "·Ì” ·Â ”Ì—Ì«·"
            .ItemData(2) = 2
        End With

 

        With Me.CboItemType
            .Clear
            .AddItem "”·⁄…"
            .AddItem "Œœ„…"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboGuar
            .Clear
            .AddItem "·Â ÷„«‰"
            .AddItem "·Ì” ·Â ÷„«‰"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboArchive
            .Clear
            .AddItem "›Ï «·√—‘Ì›"
            .AddItem "·Ì” ›Ï «·√—‘Ì›"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboAssbliedItem
            .Clear
            .AddItem "’‰› „Ã„⁄"
            .AddItem "’‰› ⁄«œÏ"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboAttachedItem
            .Clear
            .AddItem "‰⁄„"
            .AddItem "·«"
         
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        
        For i = 0 To CboItemCodeSearch.count - 1
            With Me.CboItemCodeSearch(i)
                 .Clear
                .AddItem "Typical Search"
                .AddItem "From The Start"
                .AddItem "From The End"
                .AddItem "Any Where"
            End With
            With Me.CboNameSearch(i)
                .Clear
                .AddItem "Start Name"
                .AddItem "Any Part of Name"
                End With
            Next
     

        With Me.CboSerial
            .Clear
            .AddItem "All"
            .ItemData(0) = 0
            .AddItem "Has Serial"
            .ItemData(1) = 1
            .AddItem "NO Serial"
            .ItemData(2) = 2
        End With

      

        With Me.CboItemType
            .Clear
            .AddItem "Goods"
            .AddItem "Services"
            .AddItem "All"
        End With

        With Me.CboGuar
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboArchive
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboAssbliedItem
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboAttachedItem
            .Clear
            .AddItem "YES"
            .AddItem "NO"
 
        End With

    End If

    CenterForm Me

    FormPostion Me, GetPostion
    Fg(Indx).WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset

    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    Set cSearchDcbo = Nothing

    FormPostion Me, SavePostion
    Set m_DcboItems = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer

    On Error GoTo ErrTrap

    Dim mTableName As String
    Select Case Indx
    Case 0
        mTableName = "ExpensesType"
    Case 1
        mTableName = "BanksData"
    Case 2
        mTableName = "tblBoxesData"
    Case 3
        mTableName = "TblRevenuesTypes"
    End Select

'    StrSQL = "Select * From ExpensesType "
'
    
'    StrSQL = StrSQL + " Where id <> 0 "

StrSQL = "SELECT     TT.ID, TT.Name, TT.Remarks, TT.Account_Code,   ISNULL(a2.Account_Name,TT.parent_account) parent_account , "
StrSQL = StrSQL + "                       TT.Namee, TT.TypicalProduction, IndirectCosts, ManualEntrty,"
StrSQL = StrSQL + "                       dbo.ACCOUNTS.account_serial"
StrSQL = StrSQL + "  FROM         dbo." & mTableName & " TT LEFT OUTER JOIN"
StrSQL = StrSQL + "                        dbo.ACCOUNTS ON TT.Account_Code = dbo.ACCOUNTS.Account_Code"
StrSQL = StrSQL + "                                  LEFT OUTER JOIN ACCOUNTS AS a2"
StrSQL = StrSQL + "                                           ON a2.Account_Code = TT.parent_account"

StrSQL = StrSQL + "  Where (TT.id <> 0)"
    
    If (Me.TxtAccount_Serial(Indx).Text) <> "" Then
        StrSQL = StrSQL + " AND ACCOUNTS.Account_Serial  like '%" & (Me.TxtItemID(Indx).Text) & "%'"
    End If
    

    If val(Me.TxtItemID(Indx).Text) <> 0 Then
        StrSQL = StrSQL + " AND id =" & val(Me.TxtItemID(Indx).Text)
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    If Trim(Me.TxtItemName(Indx).Text) <> "" Then
        If Me.CboNameSearch(Indx).ListIndex = 0 Then
            StrWhere = StrWhere + " and name Like '" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        ElseIf (Me.CboNameSearch(Indx).ListIndex = 1 Or Me.CboNameSearch(Indx).ListIndex = -1) Then
            StrWhere = StrWhere + " and name like '%" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        End If
    End If
Else
If Trim(Me.TxtItemName(Indx).Text) <> "" Then
        If Me.CboNameSearch(Indx).ListIndex = 0 Then
            StrWhere = StrWhere + " and namee Like '" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        ElseIf (Me.CboNameSearch(Indx).ListIndex = 1 Or Me.CboNameSearch(Indx).ListIndex = -1) Then
            StrWhere = StrWhere + " and namee like '%" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        End If
    End If

End If
    If Me.DCboGroupName(Indx).Text <> "" Then
        StrWhere = StrWhere + " and parent_account like '%" & Me.DCboGroupName(Indx).Text & "%'"
    End If

    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function



Private Function Build_SqlBox()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer

    On Error GoTo ErrTrap

    Dim mTableName As String
    Select Case Indx
    Case 0
        mTableName = "ExpensesType"
    Case 1
        mTableName = "BanksData"
    Case 2
        mTableName = "tblBoxesData"
    Case 3
        mTableName = "TblRevenuesTypes"
    End Select

'    StrSQL = "Select * From ExpensesType "
'
    
'    StrSQL = StrSQL + " Where id <> 0 "

StrSQL = "SELECT     TT.BoxID ID, TT.BoxName Name, TT.Comments Remarks, TT.Account_Code, TT.parent_account, "
StrSQL = StrSQL + "                       TT.BoxNamee Namee, "
StrSQL = StrSQL + "                       dbo.ACCOUNTS.account_serial"
StrSQL = StrSQL + "  FROM         dbo." & mTableName & " TT LEFT OUTER JOIN"
StrSQL = StrSQL + "                        dbo.ACCOUNTS ON TT.Account_Code = dbo.ACCOUNTS.Account_Code"
StrSQL = StrSQL + "  Where (TT.BoxID <> 0)"
    
    If Me.DCboGroupName(2).Text <> "" Then
        StrWhere = StrWhere + " and parent_account like '%" & Me.DCboGroupName(2).Text & "%'"
    End If
     If Me.Dcbranch(1).Text <> "" Then
        StrWhere = StrWhere + " and BranchId = " & Me.Dcbranch(1).BoundText
    End If
    
    If Me.Dcemp.Text <> "" Then
        StrWhere = StrWhere + " and EmpID = " & Me.Dcemp.BoundText
    End If



    If val(Me.TxtItemID(Indx).Text) <> 0 Then
        StrSQL = StrSQL + " AND BoxID =" & val(Me.TxtItemID(Indx).Text)
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    If Trim(Me.TxtItemName(Indx).Text) <> "" Then
        If Me.CboNameSearch(Indx).ListIndex = 0 Then
            StrWhere = StrWhere + " and Boxname Like '" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        ElseIf (Me.CboNameSearch(Indx).ListIndex = 1 Or Me.CboNameSearch(Indx).ListIndex = -1) Then
            StrWhere = StrWhere + " and Boxname like '%" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        End If
    End If
Else
If Trim(Me.TxtItemName(Indx).Text) <> "" Then
        If Me.CboNameSearch(Indx).ListIndex = 0 Then
            StrWhere = StrWhere + " and Boxnamee Like '" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        ElseIf (Me.CboNameSearch(Indx).ListIndex = 1 Or Me.CboNameSearch(Indx).ListIndex = -1) Then
            StrWhere = StrWhere + " and Boxnamee like '%" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        End If
    End If

End If
    If Me.DCboGroupName(Indx).Text <> "" Then
        StrWhere = StrWhere + " and parent_account like '%" & Me.DCboGroupName(Indx).Text & "%'"
    End If

    Build_SqlBox = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function



Private Function Build_SqlBank()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer

    On Error GoTo ErrTrap

    Dim mTableName As String
    Select Case Indx
    Case 0
        mTableName = "ExpensesType"
    Case 1
        mTableName = "BanksData"
    Case 2
        mTableName = "tblBoxesData"
    Case 3
        mTableName = "TblRevenuesTypes"
    End Select

'    StrSQL = "Select * From ExpensesType "
'
    
'    StrSQL = StrSQL + " Where id <> 0 "

StrSQL = "SELECT     TT.BankID Id, TT.BankName Name, TT.Remarks, TT.Account_Code, TT.parent_account,TT.account_no account_serial, "
StrSQL = StrSQL + "                       TT.BankNamee Namee,TT.IBan,TT.AccountName"
'StrSQL = StrSQL + "                       dbo.ACCOUNTS.account_serial"
StrSQL = StrSQL + "  FROM         dbo." & mTableName & " TT LEFT OUTER JOIN"
StrSQL = StrSQL + "                        dbo.ACCOUNTS ON TT.Account_Code = dbo.ACCOUNTS.Account_Code"
StrSQL = StrSQL + "  Where (TT.Bankid <> 0)"
    
    If (Me.TxtAccount_Serial(Indx).Text) <> "" Then
        StrSQL = StrSQL + " AND TT.account_no  like '%" & (Me.TxtAccount_Serial(Indx).Text) & "%'"
    End If

    If (Me.tXTAccount_SerialName.Text) <> "" Then
        StrSQL = StrSQL + " AND TT.AccountName  like '%" & (Me.tXTAccount_SerialName.Text) & "%'"
    End If
    
    

    If val(Me.TxtItemID(Indx).Text) <> 0 Then
        StrSQL = StrSQL + " AND Bankid =" & val(Me.TxtItemID(Indx).Text)
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    If Trim(Me.TxtItemName(Indx).Text) <> "" Then
        If Me.CboNameSearch(Indx).ListIndex = 0 Then
            StrWhere = StrWhere + " and Bankname Like '" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        ElseIf (Me.CboNameSearch(Indx).ListIndex = 1 Or Me.CboNameSearch(Indx).ListIndex = -1) Then
            StrWhere = StrWhere + " and Bankname like '%" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        End If
    End If
Else
    If Trim(Me.TxtItemName(Indx).Text) <> "" Then
        If Me.CboNameSearch(Indx).ListIndex = 0 Then
            StrWhere = StrWhere + " and Banknamee Like '" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        ElseIf (Me.CboNameSearch(Indx).ListIndex = 1 Or Me.CboNameSearch(Indx).ListIndex = -1) Then
            StrWhere = StrWhere + " and Banknamee like '%" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        End If
    End If

End If


    If Me.Dcbranch(0).Text <> "" Then
        StrWhere = StrWhere + " and BranchId = " & Me.Dcbranch(0).BoundText
    End If

    If Me.DcCurrency.Text <> "" Then
        StrWhere = StrWhere + " and Currency_ID = " & Me.DcCurrency.BoundText
    End If

 
    Build_SqlBank = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function


Private Function Build_SqlRevenues()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer

    On Error GoTo ErrTrap

    Dim mTableName As String
    Select Case Indx
    Case 0
        mTableName = "ExpensesType"
    Case 1
        mTableName = "BanksData"
    Case 2
        mTableName = "tblBoxesData"
    Case 3
        mTableName = "TblRevenuesTypes"
    End Select

'    StrSQL = "Select * From ExpensesType "
'
    
'    StrSQL = StrSQL + " Where id <> 0 "

StrSQL = "SELECT     TT.RevenuesID ID, TT.RevenuesName Name, TT.Remarks, TT.Account_Code, TT.parent_account, "
StrSQL = StrSQL + "                       TT.RevenuesNamee Namee, "
StrSQL = StrSQL + "                       dbo.ACCOUNTS.account_serial"
StrSQL = StrSQL + "  FROM         dbo." & mTableName & " TT LEFT OUTER JOIN"
StrSQL = StrSQL + "                        dbo.ACCOUNTS ON TT.Account_Code = dbo.ACCOUNTS.Account_Code"
StrSQL = StrSQL + "  Where (TT.RevenuesID <> 0)"
    
'    If (Me.tXTAccount_Serial(Indx).Text) <> "" Then
'        StrSQL = StrSQL + " AND ACCOUNTS.Account_Serial  like '%" & (Me.TxtItemID(Indx).Text) & "%'"
'    End If
    

    If val(Me.TxtItemID(Indx).Text) <> 0 Then
        StrSQL = StrSQL + " AND RevenuesID =" & val(Me.TxtItemID(Indx).Text)
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    If Trim(Me.TxtItemName(Indx).Text) <> "" Then
        If Me.CboNameSearch(Indx).ListIndex = 0 Then
            StrWhere = StrWhere + " and Revenuesname Like '" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        ElseIf (Me.CboNameSearch(Indx).ListIndex = 1 Or Me.CboNameSearch(Indx).ListIndex = -1) Then
            StrWhere = StrWhere + " and Revenuesname like '%" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        End If
    End If
Else
If Trim(Me.TxtItemName(Indx).Text) <> "" Then
        If Me.CboNameSearch(Indx).ListIndex = 0 Then
            StrWhere = StrWhere + " and Revenuesnamee Like '" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        ElseIf (Me.CboNameSearch(Indx).ListIndex = 1 Or Me.CboNameSearch(Indx).ListIndex = -1) Then
            StrWhere = StrWhere + " and Revenuesnamee like '%" & Trim(Me.TxtItemName(Indx).Text) & "%'"
        End If
    End If

End If
    If Me.DboParentAccount.Text <> "" Then
        StrWhere = StrWhere + " and parent_account like '%" & Me.DboParentAccount.Text & "%'"
    End If
    
    Build_SqlRevenues = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function


Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is Fg Then
            If Not Fg(Indx).TextMatrix(Fg(Indx).Row, 1) = "" Then
                fg_Click Indx
                Unload Me
            End If

        Else
            Cmd_Click (0)
        End If
    End If

    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Public Property Get DcboItems() As DataCombo
    Set DcboItems = m_DcboItems
End Property

Public Property Set DcboItems(ByVal vNewValue As DataCombo)
    Set m_DcboItems = vNewValue
End Property

Public Property Get RetrunType() As Integer
    RetrunType = m_RetrunType
End Property

Public Property Let RetrunType(ByVal vNewValue As Integer)
    m_RetrunType = vNewValue
    ' 0 = Retrun in the Items Screen
    ' 1 = Retrun in the Data Combo
End Property

Private Sub ChangeLang()
    Me.Caption = "Search For Expenses"
    lbl(0).Caption = " Code"
    lbl(1).Caption = " Name"
    lbl(2).Caption = "Serial Type"
    lbl(3).Caption = "Group Name"
    lbl(4).Caption = "Item ID"
    lbl(5).Caption = "Item Type"
    lbl(6).Caption = "Assembled"
    lbl(7).Caption = "Attached"
    lbl(8).Caption = "Match Type"
    lbl(9).Caption = "Guarantee"
    lbl(10).Caption = "Archives"
    lbl(11).Caption = "Match Type"
lbl(12).Caption = "Account_Serial"

    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.Fg(Indx)
    'Account_Serial
    .TextMatrix(0, .ColIndex("Account_Serial")) = "Account Serial"
        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
        .TextMatrix(0, .ColIndex("ItemNum")) = " ID"
        .TextMatrix(0, .ColIndex("KindCode")) = " Code"
        
        .TextMatrix(0, .ColIndex("KindNme")) = " Name"
        .TextMatrix(0, .ColIndex("ItemType")) = " Type"
        .TextMatrix(0, .ColIndex("HaveSerial")) = "Have Serial"
        .TextMatrix(0, .ColIndex("Parent")) = "Parent"
        .TextMatrix(0, .ColIndex("HaveGuarantee")) = "Guarantee"
        .TextMatrix(0, .ColIndex("AssbliedItem")) = "Assblied"
        .TextMatrix(0, .ColIndex("RelatedItem")) = "Attached Items"
        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub TxtItemName_Change(Index As Integer)
    
    Dim mLblIndx As Long
        If Indx = 0 Then
            mLblIndx = 8
        ElseIf Indx = 1 Then
            mLblIndx = 16
        ElseIf Indx = 2 Then
             mLblIndx = 22
        ElseIf Indx = 3 Then
             mLblIndx = 28
        ElseIf Indx = 4 Then
             mLblIndx = 23
             
        End If
    
    If Trim$(Me.TxtItemName(Index).Text) = "" Then
        Me.lbl(mLblIndx).Enabled = False
        Me.CboNameSearch(Index).Enabled = False
    Else
        Me.lbl(mLblIndx).Enabled = True
        Me.CboNameSearch(Index).Enabled = True
    End If

End Sub

