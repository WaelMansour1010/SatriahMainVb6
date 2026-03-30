VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSickleave 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8070
   ClientLeft      =   4395
   ClientTop       =   1485
   ClientWidth     =   8595
   Icon            =   "FrmSickleave.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   8595
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      TabIndex        =   4
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmSickleave.frx":6852
      Left            =   15480
      List            =   "FrmSickleave.frx":6862
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   15600
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   5
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
      Visible         =   0   'False
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   15480
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   15600
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSickleave.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSickleave.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSickleave.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSickleave.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSickleave.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSickleave.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSickleave.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSickleave.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ÕœÌÀ"
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
      ButtonImage     =   "FrmSickleave.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… "
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
      ButtonImage     =   "FrmSickleave.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
      BackColor       =   14871017
      FontSize        =   9.75
      FontName        =   "Arial"
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmSickleave.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic frm_Main 
      Height          =   8070
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   8595
      _cx             =   15161
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
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   780
         Left            =   13440
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   11760
         Begin VB.TextBox tXTRootAccount 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3240
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   780
         Index           =   18
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   8565
         _cx             =   15108
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
         BackColor       =   16777215
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
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   240
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSickleave.frx":15BA9
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   435
            TabIndex        =   17
            Top             =   240
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSickleave.frx":15F43
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   855
            TabIndex        =   18
            Top             =   240
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSickleave.frx":162DD
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   1245
            TabIndex        =   19
            Top             =   240
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   16777215
            FontSize        =   12
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmSickleave.frx":16677
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   615
            Left            =   7890
            Picture         =   "FrmSickleave.frx":16A11
            Stretch         =   -1  'True
            Top             =   120
            Width           =   525
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈⁄œ«œ«  «·«Ã«“«  «·„—÷Ì…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   3615
            TabIndex        =   20
            Top             =   240
            Width           =   2985
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   4395
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2520
         Width           =   8565
         _cx             =   15108
         _cy             =   7752
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
         Begin VSFlex8Ctl.VSFlexGrid Fg1 
            Height          =   3705
            Left            =   240
            TabIndex        =   22
            Top             =   120
            Width           =   8235
            _cx             =   14526
            _cy             =   6535
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
            FocusRect       =   0
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmSickleave.frx":17E16
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
            ComboSearch     =   0
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
            Index           =   0
            Left            =   7200
            TabIndex        =   23
            Top             =   3960
            Width           =   1050
            _ExtentX        =   1852
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
            ButtonImage     =   "FrmSickleave.frx":17EAC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   1
            Left            =   5640
            TabIndex        =   24
            Top             =   3960
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   476
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
            ButtonImage     =   "FrmSickleave.frx":18446
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   1695
         Left            =   0
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   720
         Width           =   8565
         _cx             =   15108
         _cy             =   2990
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
         Begin VB.TextBox TxtNoDay 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox TxtPeriod 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   120
            Width           =   855
         End
         Begin VB.ComboBox DcbTypePeriod 
            Height          =   315
            ItemData        =   "FrmSickleave.frx":189E0
            Left            =   165
            List            =   "FrmSickleave.frx":189EA
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox TxtNameE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   165
            TabIndex        =   43
            Top             =   1050
            Width           =   7245
         End
         Begin VB.TextBox TxtName1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   165
            TabIndex        =   41
            Top             =   600
            Width           =   7245
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   5565
            TabIndex        =   26
            Top             =   120
            Width           =   1845
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«ﬁ’Ï „œ…"
            Height          =   255
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «‰Ã·Ì“Ì"
            Height          =   315
            Index           =   1
            Left            =   7500
            TabIndex        =   44
            Top             =   1080
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ ⁄—»Ì"
            Height          =   315
            Index           =   0
            Left            =   7545
            TabIndex        =   42
            Top             =   630
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   315
            Index           =   4
            Left            =   7590
            TabIndex        =   27
            Top             =   150
            Width           =   885
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1110
         Left            =   0
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   6975
         Width           =   8565
         _cx             =   15108
         _cy             =   1958
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
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   7530
            TabIndex        =   29
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   600
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSickleave.frx":189F8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   6165
            TabIndex        =   30
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   600
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSickleave.frx":1F25A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   4740
            TabIndex        =   31
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   600
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSickleave.frx":25ABC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   3180
            TabIndex        =   32
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   600
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSickleave.frx":25E56
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   1710
            TabIndex        =   33
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   600
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSickleave.frx":261F0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   240
            TabIndex        =   34
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   600
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   582
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
            ButtonImage     =   "FrmSickleave.frx":2678A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   5040
            TabIndex        =   35
            Top             =   90
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   345
            Index           =   14
            Left            =   7350
            TabIndex        =   40
            Top             =   90
            Width           =   1140
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   195
            Index           =   0
            Left            =   3255
            TabIndex        =   39
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   38
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2370
            TabIndex        =   37
            Top             =   240
            Width           =   780
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   315
            TabIndex        =   36
            Top             =   240
            Width           =   630
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·„” Œœ„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   15480
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmSickleave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim II As Long


Private Sub Cmd_Click(Index As Integer)
Dim i As Integer
Dim k As Integer
If Me.TxtModFlg.Text <> "R" Then
Select Case Index
Case 0
RemoveGridRow

Case 1
     Fg1.Clear flexClearScrollable, flexClearEverything
                  Fg1.Rows = 1
           
End Select
End If
End Sub


Private Sub RemoveGridRow()
    With Me.Fg1
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
    ReLineGrid
End Sub
Sub ReLineGrid()
Dim i As Integer
Dim Conter As Integer
Conter = 0
With Fg1
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("FromNo"))) <> 0 Then
Conter = Conter + 1
.TextMatrix(i, .ColIndex("Serial")) = Conter
End If
Next i
End With
End Sub

Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblSickleave", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub


Private Sub DcbTypePeriod_Change()
If Me.TxtModFlg.Text <> "R" Then
If val(Me.DcbTypePeriod.ListIndex) = 0 Then
Me.TxtNoDay.Text = val(Me.TxtPeriod.Text)
ElseIf val(Me.DcbTypePeriod.ListIndex) = 1 Then
Me.TxtNoDay.Text = val(Me.TxtPeriod.Text) * 30
ElseIf val(Me.DcbTypePeriod.ListIndex) = 2 Then
Me.TxtNoDay.Text = val(Me.TxtPeriod.Text) * 360
Else
Me.TxtNoDay.Text = 0
End If
End If
End Sub

Private Sub DcbTypePeriod_Click()
DcbTypePeriod_Change
End Sub



Private Sub Fg1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'With Fg1
'If val(.TextMatrix(Row, .ColIndex("FromNo"))) <> 0 Then
'        If Row = .Rows - 1 Then
'            .Rows = .Rows + 1
'        End If
'End If
'End With

    On Error GoTo EH
    If Col = Fg1.ColIndex("ToNo") Then
        
        ZIGZAG Fg1, "FromNo", "ToNo"
    End If
    Exit Sub
EH:
    Exit Sub

End Sub
Public Sub GridKeyDown(G As Object, _
                       ByRef KeyCode As Integer, _
                       Shift As Integer, _
                       Optional ByVal CancelInsert As Boolean, _
                       Optional ByVal CancelDelete As Boolean, _
                       Optional SerStartRow As Single = 1, _
                        Optional colStartRow As String = "")
            
        Dim FirstVisibleCol As Long
        Dim LastVisibleCol As Long
        Dim FirstVisibleRow As Long
        Dim LastVisibleRow As Long
        Dim NextVisibleCol As Long
        Dim NextVisibleRow As Long
        
        
        Dim x As Integer
        
         If G.Row <= 0 Or G.Col < 0 Then Exit Sub
    ' ************************
    If SerStartRow < G.FixedRows + 1 Then
        SerStartRow = G.FixedRows
    End If
    ' ************************
    FirstVisibleRow = 0
    For x = G.FixedRows To G.Rows - 1
        If Not G.RowHidden(x) Then FirstVisibleRow = x: Exit For
    Next
    ' ************************
    LastVisibleRow = 0
    For x = G.Rows - 1 To G.FixedRows Step -1
        If Not G.RowHidden(x) Then LastVisibleRow = x: Exit For
    Next
    ' ************************
    FirstVisibleCol = 0
    For x = 1 To G.Cols - 1
        If Not G.ColHidden(x) Then FirstVisibleCol = x: Exit For
    Next
    ' ************************
    LastVisibleCol = 0
    For x = G.Cols - 1 To 1 Step -1
        If Not G.ColHidden(x) Then LastVisibleCol = x: Exit For
    Next
            Select Case KeyCode
            Case vbKeyDelete
                If Shift = 1 Then
                    If CancelDelete Then Exit Sub
                    ' ***************************
                    If LastVisibleRow >= SerStartRow + 1 Then    'G.Rows >= 3 Then
                        If G.Row < LastVisibleRow Then    'G.Rows - 1 Then
                            G.RemoveItem G.Row
                        Else
                            G.RemoveItem G.Row
                            ' ************************
                            LastVisibleRow = 0
                            For x = G.Rows - 1 To 1 Step -1
                                If Not G.RowHidden(x) Then LastVisibleRow = x: Exit For
                            Next
                            ' ************************
                            G.Row = LastVisibleRow    'G.Rows - 1
                        End If
                        GridSerial G, , SerStartRow
                    Else
                        G.Cell(flexcpText, SerStartRow, 1, SerStartRow, G.Cols - 1) = ""
                    End If
                    GridSerial G, , SerStartRow
                End If
            Case vbKeyReturn
                If G.Row = LastVisibleRow Then    'G.Rows - 1 Then ' BottomRow Then
                    If colStartRow <> "" Then
        
                        If LastVisibleCol = G.ColIndex(colStartRow) Then
                            FirstVisibleCol = G.ColIndex(colStartRow)
                        End If
                    End If

                    If G.Col = LastVisibleCol Then    'G.Cols - 1 Then
        
                        If CancelInsert Then G.FinishEditing False: GoTo mExit
                        ' ***************************
                        G.AddItem "", G.Rows
                        ' *****************
                        GridSerial G, , SerStartRow
                        G.Row = G.Rows - 1
                        G.Col = FirstVisibleCol
                    Else
                        NextVisibleCol = G.Col
                        For x = G.Col + 1 To G.Cols - 1
                            If Not G.ColHidden(x) Then NextVisibleCol = x: Exit For
                        Next
                        ' ************************
                        G.Col = NextVisibleCol
                    End If
            Else
                If G.Col = LastVisibleCol Then    'G.Cols - 1 Then
                    G.Col = FirstVisibleCol    '1
                    NextVisibleRow = G.Row
                    For x = G.Row + 1 To G.Rows - 1
                        If Not G.RowHidden(x) Then NextVisibleRow = x: Exit For
                    Next
                    ' ************************
                    G.Row = NextVisibleRow
                Else
                    NextVisibleCol = G.Col
                    For x = G.Col + 1 To G.Cols - 1
                        If Not G.ColHidden(x) Then NextVisibleCol = x: Exit For
                    Next
                    ' ************************
                    G.Col = NextVisibleCol
                End If
            End If
mExit:
        KeyCode = 0
        End Select

        Exit Sub
End Sub
Public Sub GridSerial(vsGrd As Object, _
                      Optional chkRowHidden As Boolean = False, _
                      Optional SerStartRow As Single = 0)
    Dim j As Long
    Dim ss As Long
    If SerStartRow = 0 Then SerStartRow = vsGrd.FixedRows

    For ss = SerStartRow To vsGrd.Rows - 1
        If chkRowHidden Then
            If Not vsGrd.RowHidden(ss) Then j = j + 1
        Else
            j = j + 1
        End If
        vsGrd.TextMatrix(ss, 0) = j
    Next

End Sub

Sub ZIGZAG(myGrid As Object, _
           strStartCol As String, _
           strEndCol As String, _
           Optional IsCountinous As Boolean = True, Optional mRow As Long = 0)
    Dim ArabicInterface As String
    Dim i As Integer
    ArabicInterface = IIf(SystemOptions.UserInterface = EnglishInterface, False, True)
    IsCountinous = False
    If mRow = 0 Then mRow = myGrid.Row
    If myGrid.Rows - 1 = mRow Then myGrid.Rows = myGrid.Rows + 1
    If val(myGrid.TextMatrix(mRow, myGrid.ColIndex(strEndCol))) = 0 Then
        Exit Sub
    End If
    If Not IsNumeric(myGrid.TextMatrix(mRow, myGrid.ColIndex(strStartCol))) Then
        If IsCountinous Then
            MsgBox IIf(ArabicInterface, "ÌÃ» √‰ ÌﬂÊ‰ «·—ﬁ„ «·√Ê· —ﬁ„", "The First Number must be a digit")
        End If
        myGrid.TextMatrix(mRow, myGrid.ColIndex(strStartCol)) = "0"
    ElseIf val(myGrid.TextMatrix(mRow, myGrid.ColIndex(strStartCol))) < 0 Then
        If IsCountinous Then
            MsgBox IIf(ArabicInterface, "ÌÃ» √·« ÌﬂÊ‰ «·—ﬁ„ «·√Ê· √ﬁ· „‰ «·’›—", "the first number must be More than Zero")
        End If
        myGrid.TextMatrix(mRow, myGrid.ColIndex(strStartCol)) = "0"
    ElseIf Not IsNumeric(myGrid.TextMatrix(mRow, myGrid.ColIndex(strEndCol))) Then
        If IsCountinous Then
            MsgBox IIf(ArabicInterface, "ÌÃ» √‰ ÌﬂÊ‰ «·—ﬁ„ «·À«‰Ì —ﬁ„", "The Second Number must be a Digit")
        End If
        myGrid.TextMatrix(mRow, myGrid.ColIndex(strEndCol)) = myGrid.TextMatrix(mRow, myGrid.ColIndex(strStartCol))
    ElseIf val(myGrid.TextMatrix(mRow, myGrid.ColIndex(strEndCol))) < 0 Then
        If IsCountinous Then
            MsgBox IIf(ArabicInterface, "ÌÃ» √·« ÌﬂÊ‰ «·—ﬁ„ «·À«‰Ì √ﬁ· „‰ «·’›—", "The Second Number must not be Less than Zero")
        End If
        myGrid.TextMatrix(mRow, myGrid.ColIndex(strEndCol)) = myGrid.TextMatrix(mRow, myGrid.ColIndex(strStartCol))
    ElseIf val(myGrid.TextMatrix(mRow, myGrid.ColIndex(strEndCol))) < val(myGrid.TextMatrix(mRow, myGrid.ColIndex(strStartCol))) Then
        If IsCountinous Then
            MsgBox IIf(ArabicInterface, "ÌÃ» √·« ÌﬂÊ‰ «·—ﬁ„ «·À«‰Ì √ﬁ· „‰ «·—ﬁ„ «·√Ê·", "The Second Number must not be smaller than The First Number")
        End If
        myGrid.TextMatrix(mRow, myGrid.ColIndex(strEndCol)) = myGrid.TextMatrix(mRow, myGrid.ColIndex(strStartCol))
    ElseIf val(myGrid.TextMatrix(mRow, myGrid.ColIndex(strEndCol))) = val(myGrid.TextMatrix(mRow - 1, myGrid.ColIndex(strEndCol))) Then
        If IsCountinous Then
            MsgBox IIf(ArabicInterface, "ÌÃ» ⁄œ„  ﬂ—«— «·› —« ", "Can't Repeat Periods")
        End If
        myGrid.TextMatrix(mRow, myGrid.ColIndex(strEndCol)) = myGrid.TextMatrix(mRow, myGrid.ColIndex(strStartCol)) + 1
    End If
    
    For i = 2 To myGrid.Rows - 1
       ' If myGrid.TextMatrix(i, myGrid.ColIndex(strStartCol)) <> "" Then
            If IsCountinous Then
                myGrid.TextMatrix(i, myGrid.ColIndex(strStartCol)) = myGrid.TextMatrix(i - 1, myGrid.ColIndex(strEndCol))
            Else
                myGrid.TextMatrix(i, myGrid.ColIndex(strStartCol)) = val(myGrid.TextMatrix(i - 1, myGrid.ColIndex(strEndCol))) + 1
                If val(myGrid.TextMatrix(i, myGrid.ColIndex(strEndCol))) = 0 Or val(myGrid.TextMatrix(i, myGrid.ColIndex(strEndCol))) <= val(myGrid.TextMatrix(i, myGrid.ColIndex(strStartCol))) Then
                    myGrid.TextMatrix(i, myGrid.ColIndex(strEndCol)) = val(myGrid.TextMatrix(i, myGrid.ColIndex(strStartCol))) + 1
                End If
            End If
       ' End If
    Next i
End Sub


Private Sub Fg1_KeyDown(KeyCode As Integer, Shift As Integer)
GridKeyDown Fg1, KeyCode, Shift, False, False, Fg1.Row
End Sub

Private Sub Fg1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
 Fg1_KeyDown KeyCode, 0
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from  TblSickleave  order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
   Dcombos.GetUsers Me.DCboUserName
   If SystemOptions.UserInterface = ArabicInterface Then
   With DcbTypePeriod
   .Clear
   .AddItem "ÌÊ„"
   .AddItem "‘Â—"
   .AddItem "”‰…"
   End With
   Else
   With DcbTypePeriod
   .Clear
   .AddItem "Day"
   .AddItem "Month"
   .AddItem "Year"
   End With
   End If
    BtnLast_Click
    ShowTip
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
   Me.Refresh
   FiLLTXT
ErrTrap:
End Sub

' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
    Dim i As Integer
      If Me.TxtModFlg.Text = "E" Then
          Cn.Execute " delete from TblSickleaveDet where SickLID=" & val(TxtSerial1.Text) & "  "
      End If
   RsSavRec.Fields("UserID").value = val(Me.DCboUserName.BoundText)
   RsSavRec.Fields("Name").value = Me.TxtName1.Text
   RsSavRec.Fields("NameE").value = Me.TxtNameE.Text
   RsSavRec.Fields("NoDay").value = val(TxtNoDay.Text)
   RsSavRec.Fields("Period").value = val(TxtPeriod.Text)
   RsSavRec.Fields("TypePeriod").value = val(Me.DcbTypePeriod.ListIndex)
   RsSavRec.update
  ''//////////////////////////

  Dim RsDevsub As ADODB.Recordset
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblSickleaveDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim str2 As String
    With Me.Fg1
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("FromNo"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("SickLID").value = val(Me.TxtSerial1.Text)
                RsDevsub("FromNo").value = IIf((.TextMatrix(i, .ColIndex("FromNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("FromNo"))))
                RsDevsub("ToNo").value = IIf((.TextMatrix(i, .ColIndex("ToNo"))) = "", Null, val(.TextMatrix(i, .ColIndex("ToNo"))))
                RsDevsub("PerstageSal").value = IIf((.TextMatrix(i, .ColIndex("PerstageSal"))) = "", Null, val(.TextMatrix(i, .ColIndex("PerstageSal"))))
       RsDevsub.update
      End If
     Next i
    End With
'''///////////////
   
    Dim Msg As String
      Select Case Me.TxtModFlg.Text
        Case "N"
            
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
                FiLLTXT
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub
Sub FullGri()
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim i As Integer
Dim sql As String
    Fg1.Clear flexClearScrollable, flexClearEverything
    Fg1.Rows = 1
sql = "SELECT   * from TblSickleaveDet where  SickLID=" & val(TxtSerial1.Text) & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With Fg1
Rs3.MoveFirst
.Rows = .Rows + Rs3.RecordCount
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("FromNo")) = IIf(IsNull(Rs3("FromNo").value), "", Rs3("FromNo").value)
.TextMatrix(i, .ColIndex("ToNo")) = IIf(IsNull(Rs3("ToNo").value), "", Rs3("ToNo").value)
.TextMatrix(i, .ColIndex("PerstageSal")) = IIf(IsNull(Rs3("PerstageSal").value), "", Rs3("PerstageSal").value)
Rs3.MoveNext
Next i
End With
End If
End Sub
' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()

   On Error GoTo ErrTrap
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbTypePeriod.ListIndex = IIf(IsNull(RsSavRec.Fields("TypePeriod").value), -1, RsSavRec.Fields("TypePeriod").value)
    TxtPeriod.Text = IIf(IsNull(RsSavRec.Fields("Period").value), "", RsSavRec.Fields("Period").value)
    TxtNoDay.Text = IIf(IsNull(RsSavRec.Fields("NoDay").value), 0, RsSavRec.Fields("NoDay").value)
    TxtName1.Text = IIf(IsNull(RsSavRec.Fields("Name").value), "", RsSavRec.Fields("Name").value)
    TxtNameE.Text = IIf(IsNull(RsSavRec.Fields("NameE").value), "", RsSavRec.Fields("NameE").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGri
ErrTrap:
End Sub

' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim Total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    If val(TxtPeriod.Text) = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «œŒ«· «ﬁ’Ï „œ… "
    Else
    MsgBox "Please Eneter Maximum Duration"
    End If
    TxtPeriod.SetFocus
    Exit Sub
    End If
    If val(Me.DcbTypePeriod.ListIndex) = -1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «ﬁ’Ï „œ… "
    Else
    MsgBox "Please Select Type Maximum Duration"
    End If
    DcbTypePeriod.SetFocus
    Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
If TxtName1.Text = "" Then
MsgBox "Ì—ÃÏ «œŒ«· «·«”„"
TxtName1.SetFocus
Exit Sub
End If
Else
If TxtNameE.Text = "" Then
MsgBox "Please Enter The Name"
TxtNameE.SetFocus
Exit Sub
End If
End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text
            '------------------------------ new record ----------------------------
        Case "N"
                  '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub

Private Sub TxtPeriod_Change()
DcbTypePeriod_Change
End Sub

Private Sub TxtPeriod_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPeriod.Text, 0)
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "ID=" & RecId, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
  End Function
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
End Sub
' delet sub
Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim sql As String
    Dim i As Integer
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim x As Integer
    Dim ID As Double
    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
  
    If x = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
         Cn.Execute "Delete from TblSickleaveDet where SickLID=" & val(TxtSerial1.Text) & " "
                RsSavRec.Find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.Delete
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
            Fg1.Clear flexClearScrollable, flexClearEverything
                  Fg1.Rows = 1
                 If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
     LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete the record"
            StrMSG = StrMSG & " Is related to with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           Cn.Errors.Clear
    End Select

End Sub

' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
               btnSave_Click
        Case vbCancel
              Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
    'XPDtbTrans.Enabled = True
      '  Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg.Text = "R" Then
   ' XPDtbTrans.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
    End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
  ' XPDtbTrans.Enabled = True
  '     Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    '    Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub

' move btowen recored
Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
 
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
    Fg1.Rows = Fg1.Rows + 1
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & CHR(13)
            Msg = Msg & " You can not edit this the record now" & CHR(13)
            Msg = Msg & "It was being edited by another user on the network"
           
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    clear_all Me
    TxtModFlg.Text = "N"
      Fg1.Clear flexClearScrollable, flexClearEverything
      Fg1.Rows = 2
     Me.DCboUserName.BoundText = user_id
 
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
      
        Exit Sub
    End If
BegnieWork:
     If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext
        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
       
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With
ErrTrap:
End Sub
' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If
    'End If
    Exit Sub
ErrTrap:
End Sub
Private Sub ChangeLang()
On Error GoTo ErrTrap
 Label1(2).Caption = "Sick leave Settings"
 
 lbl(4).Caption = "No."
 Label10.Caption = "Max duration"
 lbl(0).Caption = "Arabic Name"
 lbl(1).Caption = "English Name"
 
 With Fg1
    .TextMatrix(0, .ColIndex("Serial")) = "No."
    .TextMatrix(0, .ColIndex("FromNo")) = "From"
    .TextMatrix(0, .ColIndex("ToNo")) = "To"
    .TextMatrix(0, .ColIndex("PerstageSal")) = "Salary Percentage"
 End With
 
 Cmd(0).Caption = "Delete"
 Cmd(1).Caption = "Delete all"
 
 lbl(14).Caption = "Edited by"
 Label2(0).Caption = "current record"
 Label2(1).Caption = "Record Count"
 
 btnNew.Caption = "New"
 btnModify.Caption = "Edit"
 btnSave.Caption = "Save"
 BtnUndo.Caption = "Undo"
 btnDelete.Caption = "Delete"
 btnCancel.Caption = "Exit"
 
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblSickleave"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

