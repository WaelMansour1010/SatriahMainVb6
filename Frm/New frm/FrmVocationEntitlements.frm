VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmVocationEntitlements 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„” Õﬁ«  «·ﬁÌ«„ »«Ã«“…"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18420
   Icon            =   "FrmVocationEntitlements.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10635
   ScaleWidth      =   18420
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   2895
      Left            =   20160
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   16335
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Ì«„ «·€Ì«»"
         Height          =   735
         Left            =   16440
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1680
         Width           =   2265
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—« »"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   32
            Left            =   1080
            TabIndex        =   21
            Top             =   0
            Width           =   525
         End
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·„” Õﬁ« "
         Height          =   195
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   -360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· ⁄Âœ"
         Height          =   195
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   -360
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   18690
      TabIndex        =   11
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox ¡¡ 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   19110
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   18990
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   18870
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox xxxxxx 
      Height          =   285
      Left            =   18990
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   5535
      Left            =   18360
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   16320
      _cx             =   28787
      _cy             =   9763
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
      Caption         =   "„” Õﬁ«  «·«Ã«“Â|Õ«·Â «·«⁄ „«œ"
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
      Picture(0)      =   "FrmVocationEntitlements.frx":038A
      Flags(1)        =   2
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5070
         Left            =   16965
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   45
         Width           =   16230
         _cx             =   28628
         _cy             =   8943
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
            TabIndex        =   3
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
            FormatString    =   $"FrmVocationEntitlements.frx":0724
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
            TabIndex        =   6
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
            TabIndex        =   4
            Top             =   4560
            Width           =   3375
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   5070
         Index           =   15
         Left            =   45
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   45
         Width           =   16230
         _cx             =   28628
         _cy             =   8943
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
         _GridInfo       =   $"FrmVocationEntitlements.frx":0870
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5040
            Index           =   16
            Left            =   15
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   15
            Width           =   16200
            _cx             =   28575
            _cy             =   8890
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   2910
               Index           =   62
               Left            =   3090
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   1335
               Width           =   735
            End
         End
      End
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   18870
      TabIndex        =   12
      Top             =   3570
      Visible         =   0   'False
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
      Left            =   19230
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
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
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   10635
      Index           =   6
      Left            =   0
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   18420
      _cx             =   32491
      _cy             =   18759
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   5070
         Left            =   0
         TabIndex        =   147
         Top             =   4035
         Width           =   18420
         _cx             =   32491
         _cy             =   8943
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
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   14871017
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "»Ì«‰«  «”«”Ì…|«·«⁄ „«œ"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   4695
            Left            =   45
            TabIndex        =   148
            TabStop         =   0   'False
            Top             =   45
            Width           =   18330
            _cx             =   32332
            _cy             =   8281
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   3045
               Index           =   9
               Left            =   0
               TabIndex        =   149
               TabStop         =   0   'False
               Top             =   0
               Width           =   5685
               _cx             =   10028
               _cy             =   5371
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
               Begin VB.TextBox txtPaymentRecommended 
                  Alignment       =   2  'Center
                  Height          =   375
                  Left            =   2550
                  RightToLeft     =   -1  'True
                  TabIndex        =   259
                  Top             =   2640
                  Width           =   1500
               End
               Begin VB.TextBox Total 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Enabled         =   0   'False
                  Height          =   435
                  Left            =   1080
                  TabIndex        =   159
                  Top             =   2220
                  Width           =   2970
               End
               Begin VB.TextBox TxtAdvance 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Height          =   435
                  Left            =   2835
                  TabIndex        =   158
                  Top             =   465
                  Width           =   795
               End
               Begin VB.TextBox TxtOther 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Height          =   435
                  Left            =   405
                  TabIndex        =   157
                  Top             =   465
                  Width           =   945
               End
               Begin VB.TextBox TxtTotalCut 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Enabled         =   0   'False
                  Height          =   390
                  Left            =   405
                  TabIndex        =   156
                  Top             =   1380
                  Width           =   3225
               End
               Begin VB.TextBox NetTotal 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Enabled         =   0   'False
                  Height          =   450
                  Left            =   690
                  TabIndex        =   155
                  Top             =   2565
                  Visible         =   0   'False
                  Width           =   3360
               End
               Begin VB.CheckBox ChBooked 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " „ «·ÕÃ“"
                  Height          =   330
                  Left            =   2940
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   1800
                  Width           =   960
               End
               Begin VB.CheckBox ChDelivery 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «· –«ﬂ—"
                  Height          =   330
                  Left            =   1470
                  RightToLeft     =   -1  'True
                  TabIndex        =   153
                  Top             =   1800
                  Width           =   1380
               End
               Begin VB.TextBox TxtValueTickt 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0C0&
                  Height          =   435
                  Left            =   405
                  TabIndex        =   152
                  Top             =   1800
                  Width           =   945
               End
               Begin VB.TextBox TxtDecrease 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Height          =   435
                  Left            =   2835
                  Locked          =   -1  'True
                  TabIndex        =   151
                  Top             =   900
                  Width           =   795
               End
               Begin VB.TextBox TxtInsuranceValue 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  Height          =   450
                  Left            =   405
                  TabIndex        =   150
                  TabStop         =   0   'False
                  Top             =   915
                  Width           =   945
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   330
                  Index           =   4
                  Left            =   3660
                  TabIndex        =   160
                  Top             =   465
                  Width           =   270
                  _Version        =   786432
                  _ExtentX        =   476
                  _ExtentY        =   582
                  _StockProps     =   79
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   330
                  Index           =   5
                  Left            =   1365
                  TabIndex        =   161
                  Top             =   465
                  Width           =   255
                  _Version        =   786432
                  _ExtentX        =   450
                  _ExtentY        =   582
                  _StockProps     =   79
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   330
                  Index           =   6
                  Left            =   3660
                  TabIndex        =   162
                  Top             =   915
                  Width           =   270
                  _Version        =   786432
                  _ExtentX        =   476
                  _ExtentY        =   582
                  _StockProps     =   79
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   330
                  Index           =   7
                  Left            =   1365
                  TabIndex        =   163
                  Top             =   900
                  Width           =   255
                  _Version        =   786432
                  _ExtentX        =   450
                  _ExtentY        =   582
                  _StockProps     =   79
                  UseVisualStyle  =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Ê’Ï »”œ«œ"
                  Height          =   315
                  Index           =   75
                  Left            =   4455
                  TabIndex        =   260
                  Top             =   2700
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «· –«ﬂ—"
                  Height          =   375
                  Index           =   40
                  Left            =   1470
                  TabIndex        =   172
                  Top             =   1530
                  Visible         =   0   'False
                  Width           =   1545
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”·›… „” Õﬁ…"
                  Height          =   405
                  Index           =   26
                  Left            =   4305
                  TabIndex        =   171
                  Top             =   465
                  Width           =   1140
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ—Ï Œ’„"
                  Height          =   405
                  Index           =   29
                  Left            =   1605
                  TabIndex        =   170
                  Top             =   465
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FF80FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·«” ﬁÿ«⁄«  «·„«·Ì…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF80FF&
                  Height          =   345
                  Index           =   19
                  Left            =   825
                  TabIndex        =   169
                  Top             =   0
                  Width           =   2880
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Ã„«·Ì «·„” ﬁÿ⁄"
                  Height          =   525
                  Index           =   37
                  Left            =   3105
                  TabIndex        =   168
                  Top             =   1380
                  Width           =   2340
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«” ﬁÿ«⁄«  «·„ €Ì—…"
                  Height          =   405
                  Index           =   38
                  Left            =   3510
                  TabIndex        =   167
                  Top             =   915
                  Width           =   1935
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Êﬁ› «· –«ﬂ—"
                  Height          =   285
                  Index           =   39
                  Left            =   4065
                  TabIndex        =   166
                  Top             =   1905
                  Width           =   1380
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Ã„«·Ì «·„” Õﬁ« "
                  Height          =   390
                  Index           =   41
                  Left            =   3915
                  TabIndex        =   165
                  Top             =   2220
                  Width           =   1530
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬁÌ„… «· √„Ì‰"
                  Height          =   375
                  Index           =   57
                  Left            =   1755
                  TabIndex        =   164
                  Top             =   915
                  Width           =   990
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   3045
               Index           =   8
               Left            =   5775
               TabIndex        =   173
               TabStop         =   0   'False
               Top             =   0
               Width           =   4485
               _cx             =   7911
               _cy             =   5371
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
               Begin VB.TextBox TxtDaySalary 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFC0&
                  Height          =   390
                  Left            =   1290
                  TabIndex        =   185
                  Top             =   510
                  Width           =   780
               End
               Begin VB.TextBox TxtSalary 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFC0&
                  Height          =   390
                  Left            =   285
                  TabIndex        =   184
                  Top             =   525
                  Width           =   870
               End
               Begin VB.TextBox TxtDayIncrease 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFC0&
                  Height          =   375
                  Left            =   3720
                  TabIndex        =   183
                  Top             =   1065
                  Visible         =   0   'False
                  Width           =   750
               End
               Begin VB.TextBox TxtIncrease 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFC0&
                  Height          =   390
                  Left            =   285
                  Locked          =   -1  'True
                  TabIndex        =   182
                  Top             =   930
                  Width           =   1830
               End
               Begin VB.TextBox TxtDaySalVocation 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFC0&
                  Height          =   375
                  Left            =   1290
                  TabIndex        =   181
                  Top             =   1335
                  Width           =   780
               End
               Begin VB.TextBox TxtSalaryVocation 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFC0&
                  Height          =   375
                  Left            =   285
                  TabIndex        =   180
                  Top             =   1335
                  Width           =   870
               End
               Begin VB.TextBox TxtDayEntitOther 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFC0&
                  Height          =   375
                  Left            =   4245
                  TabIndex        =   179
                  Top             =   1860
                  Visible         =   0   'False
                  Width           =   765
               End
               Begin VB.TextBox TxtSalEntitOther 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFC0&
                  Height          =   360
                  Left            =   285
                  TabIndex        =   178
                  Top             =   2145
                  Width           =   1830
               End
               Begin VB.TextBox TxtTolaMostak 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFC0&
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   285
                  TabIndex        =   177
                  Top             =   2535
                  Width           =   1830
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "⁄—÷"
                  Height          =   345
                  Left            =   3225
                  RightToLeft     =   -1  'True
                  TabIndex        =   176
                  Top             =   150
                  Width           =   900
               End
               Begin VB.TextBox TxtPreSalary 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFC0&
                  Height          =   420
                  Left            =   810
                  Locked          =   -1  'True
                  TabIndex        =   175
                  Top             =   1680
                  Width           =   1305
               End
               Begin VB.CommandButton Command3 
                  Caption         =   "⁄—÷"
                  Height          =   375
                  Left            =   285
                  RightToLeft     =   -1  'True
                  TabIndex        =   174
                  Top             =   1680
                  Width           =   540
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   330
                  Index           =   0
                  Left            =   2130
                  TabIndex        =   186
                  Top             =   525
                  Width           =   285
                  _Version        =   786432
                  _ExtentX        =   503
                  _ExtentY        =   582
                  _StockProps     =   79
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   300
                  Index           =   1
                  Left            =   2130
                  TabIndex        =   187
                  Top             =   930
                  Width           =   285
                  _Version        =   786432
                  _ExtentX        =   503
                  _ExtentY        =   529
                  _StockProps     =   79
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   270
                  Index           =   2
                  Left            =   2130
                  TabIndex        =   188
                  Top             =   1335
                  Width           =   285
                  _Version        =   786432
                  _ExtentX        =   503
                  _ExtentY        =   476
                  _StockProps     =   79
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   315
                  Index           =   3
                  Left            =   2130
                  TabIndex        =   189
                  Top             =   2130
                  Width           =   285
                  _Version        =   786432
                  _ExtentX        =   503
                  _ExtentY        =   556
                  _StockProps     =   79
                  UseVisualStyle  =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   300
                  Index           =   8
                  Left            =   2130
                  TabIndex        =   190
                  Top             =   1680
                  Width           =   285
                  _Version        =   786432
                  _ExtentX        =   503
                  _ExtentY        =   529
                  _StockProps     =   79
                  UseVisualStyle  =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„” Õﬁ«  «·„«·Ì…"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   285
                  Index           =   14
                  Left            =   690
                  TabIndex        =   199
                  Top             =   0
                  Width           =   2580
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—« » «·‘Â— «·Õ«·Ì"
                  Height          =   375
                  Index           =   15
                  Left            =   2805
                  TabIndex        =   198
                  Top             =   525
                  Width           =   1500
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·«Ì«„"
                  Height          =   345
                  Index           =   12
                  Left            =   1440
                  TabIndex        =   197
                  Top             =   255
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„»·€ «·„” Õﬁ"
                  Height          =   345
                  Index           =   13
                  Left            =   285
                  TabIndex        =   196
                  Top             =   255
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«÷«›«  «·„ €Ì—…"
                  Height          =   360
                  Index           =   16
                  Left            =   2805
                  TabIndex        =   195
                  Top             =   930
                  Width           =   1500
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—« » «·«Ã«“… «·„” Õﬁ…"
                  Height          =   330
                  Index           =   17
                  Left            =   2670
                  TabIndex        =   194
                  Top             =   1335
                  Width           =   1635
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ—Ï «÷«›…"
                  Height          =   375
                  Index           =   18
                  Left            =   2805
                  TabIndex        =   193
                  Top             =   2130
                  Width           =   1500
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Ã„«·Ì «·„»·€ «·„” Õﬁ"
                  Height          =   495
                  Index           =   34
                  Left            =   2535
                  TabIndex        =   192
                  Top             =   2535
                  Width           =   1770
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«ÃÊ— ”«»ﬁ…"
                  Height          =   360
                  Index           =   63
                  Left            =   2805
                  TabIndex        =   191
                  Top             =   1680
                  Width           =   1500
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   3045
               Index           =   7
               Left            =   10260
               TabIndex        =   200
               TabStop         =   0   'False
               Top             =   0
               Width           =   8055
               _cx             =   14208
               _cy             =   5371
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
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   2220
                  Left            =   135
                  TabIndex        =   201
                  Top             =   420
                  Width           =   7845
                  _cx             =   13838
                  _cy             =   3916
                  Appearance      =   1
                  BorderStyle     =   1
                  Enabled         =   0   'False
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
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmVocationEntitlements.frx":08A6
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  „›—œ«  «·«Ã«“…"
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Index           =   5
                  Left            =   3510
                  TabIndex        =   204
                  Top             =   150
                  Width           =   2085
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·„›—œ"
                  ForeColor       =   &H00FF0000&
                  Height          =   345
                  Index           =   11
                  Left            =   2625
                  TabIndex        =   203
                  Top             =   2655
                  Width           =   2085
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   345
                  Index           =   10
                  Left            =   405
                  TabIndex        =   202
                  Top             =   2655
                  Width           =   1005
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1800
               Index           =   12
               Left            =   -240
               TabIndex        =   209
               TabStop         =   0   'False
               Top             =   3000
               Visible         =   0   'False
               Width           =   18690
               _cx             =   32967
               _cy             =   3175
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
               Begin VSFlex8Ctl.VSFlexGrid Grid1 
                  Height          =   1245
                  Left            =   795
                  TabIndex        =   210
                  Top             =   165
                  Width           =   17385
                  _cx             =   30665
                  _cy             =   2196
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   72
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmVocationEntitlements.frx":09DE
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
               Begin XtremeSuiteControls.CheckBox Check21 
                  Height          =   375
                  Left            =   21675
                  TabIndex        =   211
                  Top             =   0
                  Width           =   1350
                  _Version        =   786432
                  _ExtentX        =   2381
                  _ExtentY        =   661
                  _StockProps     =   79
                  Caption         =   " ÕœÌœ «·ﬂ·"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   330
                  Left            =   18255
                  RightToLeft     =   -1  'True
                  TabIndex        =   212
                  Top             =   0
                  Width           =   240
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   1695
               Left            =   135
               TabIndex        =   213
               TabStop         =   0   'False
               Top             =   3000
               Width           =   18270
               _cx             =   32226
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
               BackColor       =   12648447
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
                  Height          =   1215
                  Left            =   -3960
                  TabIndex        =   214
                  Top             =   240
                  Width           =   22590
                  _cx             =   39846
                  _cy             =   2143
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   65
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmVocationEntitlements.frx":12EA
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   1455
                  Index           =   3
                  Left            =   9960
                  TabIndex        =   215
                  TabStop         =   0   'False
                  Top             =   2190
                  Width           =   2325
                  _cx             =   4101
                  _cy             =   2566
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
                     Left            =   90
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   217
                     Top             =   240
                     Width           =   1755
                  End
                  Begin VB.ComboBox CmbMonth 
                     Height          =   315
                     Left            =   90
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   216
                     Top             =   540
                     Width           =   1755
                  End
                  Begin ImpulseButton.ISButton CmdOk 
                     Height          =   315
                     Left            =   90
                     TabIndex        =   218
                     Top             =   855
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   556
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "⁄—÷  "
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
                     ButtonImage     =   "FrmVocationEntitlements.frx":1AE1
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”‰…"
                     Height          =   15
                     Index           =   52
                     Left            =   90
                     RightToLeft     =   -1  'True
                     TabIndex        =   220
                     Top             =   1815
                     Width           =   1755
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Â—"
                     Height          =   15
                     Index           =   53
                     Left            =   90
                     RightToLeft     =   -1  'True
                     TabIndex        =   219
                     Top             =   1860
                     Width           =   1755
                  End
               End
               Begin MSDataListLib.DataCombo Dcemp 
                  Height          =   315
                  Left            =   945
                  TabIndex        =   221
                  Top             =   2190
                  Width           =   3585
                  _ExtentX        =   6324
                  _ExtentY        =   556
                  _Version        =   393216
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
                  Caption         =   "„ÊŸ› „Õœœ"
                  DataField       =   "Õœœ"
                  Height          =   390
                  Index           =   54
                  Left            =   4365
                  RightToLeft     =   -1  'True
                  TabIndex        =   223
                  Top             =   2205
                  Width           =   1200
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   330
                  Left            =   17445
                  RightToLeft     =   -1  'True
                  TabIndex        =   222
                  Top             =   0
                  Width           =   600
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1770
               Index           =   10
               Left            =   135
               TabIndex        =   205
               TabStop         =   0   'False
               Top             =   3000
               Visible         =   0   'False
               Width           =   18270
               _cx             =   32226
               _cy             =   3122
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
               Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
                  Height          =   1275
                  Left            =   -4455
                  TabIndex        =   206
                  Top             =   330
                  Width           =   22590
                  _cx             =   39846
                  _cy             =   2249
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
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmVocationEntitlements.frx":1E7B
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
               Begin MSComCtl2.DTPicker DTPicker2 
                  Height          =   495
                  Left            =   10005
                  TabIndex        =   207
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   873
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   222232577
                  CurrentDate     =   38784
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   17325
                  RightToLeft     =   -1  'True
                  TabIndex        =   208
                  Top             =   0
                  Width           =   585
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   4695
            Left            =   19065
            TabIndex        =   225
            TabStop         =   0   'False
            Top             =   45
            Width           =   18330
            _cx             =   32332
            _cy             =   8281
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
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
               Height          =   4020
               Left            =   135
               TabIndex        =   226
               Tag             =   "1"
               Top             =   285
               Width           =   17745
               _cx             =   31300
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
               FormatString    =   $"FrmVocationEntitlements.frx":1FB9
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
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   330
               Left            =   11040
               RightToLeft     =   -1  'True
               TabIndex        =   228
               Top             =   4245
               Width           =   3750
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   315
               Left            =   7245
               RightToLeft     =   -1  'True
               TabIndex        =   227
               Top             =   4845
               Width           =   3675
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Width           =   18450
         _cx             =   32544
         _cy             =   1244
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   21.75
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
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "  „” Õﬁ«  «·ﬁÌ«„ »«Ã«“…  "
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
            TabIndex        =   25
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVocationEntitlements.frx":20FC
            ColorButton     =   16777215
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
            TabIndex        =   26
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVocationEntitlements.frx":2496
            ColorButton     =   16777215
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
            TabIndex        =   27
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVocationEntitlements.frx":2830
            ColorButton     =   16777215
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
            TabIndex        =   28
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVocationEntitlements.frx":2BCA
            ColorButton     =   16777215
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
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   555
            Index           =   27
            Left            =   2400
            TabIndex        =   29
            Top             =   0
            Width           =   2205
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   5880
            Picture         =   "FrmVocationEntitlements.frx":2F64
            Stretch         =   -1  'True
            Top             =   0
            Width           =   525
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3315
         Index           =   18
         Left            =   0
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   720
         Width           =   18420
         _cx             =   32491
         _cy             =   5847
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
         Begin VB.TextBox txtDaysCountPay 
            Alignment       =   2  'Center
            Height          =   390
            Left            =   5700
            TabIndex        =   264
            TabStop         =   0   'False
            Top             =   2280
            Width           =   945
         End
         Begin VB.TextBox TxtWithOutSala2 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   10710
            RightToLeft     =   -1  'True
            TabIndex        =   262
            Top             =   2850
            Width           =   780
         End
         Begin VB.TextBox TxtAddDay 
            Alignment       =   2  'Center
            Height          =   390
            Left            =   5700
            TabIndex        =   252
            TabStop         =   0   'False
            Top             =   1875
            Width           =   945
         End
         Begin VB.TextBox TxtDiscouDay 
            Alignment       =   2  'Center
            Height          =   390
            Left            =   3105
            TabIndex        =   241
            TabStop         =   0   'False
            Top             =   1875
            Width           =   975
         End
         Begin VB.TextBox IDes 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   285
            RightToLeft     =   -1  'True
            TabIndex        =   230
            Top             =   2730
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.TextBox TxtNoVaction 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   229
            Top             =   2730
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.TextBox LastBalanceMonth 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   3120
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   2400
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox TxtLastDayVoc 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   18540
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   750
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox TxtContDay 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   405
            Left            =   15510
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   2865
            Width           =   825
         End
         Begin VB.TextBox TxtWithOutSala1 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   12315
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   2865
            Width           =   780
         End
         Begin VB.TextBox TxtNewAbsent 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   13890
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   2865
            Width           =   855
         End
         Begin VB.TextBox TxtRemark 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   2865
            Width           =   3135
         End
         Begin VB.TextBox TxtToalAbsent 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   405
            Left            =   8700
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   2865
            Width           =   855
         End
         Begin VB.TextBox TxtDuVocation 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   405
            Left            =   6225
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   2865
            Width           =   960
         End
         Begin VB.TextBox TxtTotalDay 
            Alignment       =   2  'Center
            Height          =   405
            Left            =   4065
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   2865
            Width           =   1110
         End
         Begin VB.TextBox TxtGetInsurance 
            Alignment       =   2  'Center
            Height          =   390
            Left            =   0
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.TextBox TxtNoMonth 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   1905
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   585
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox XPTxtID 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   390
            Left            =   14880
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   150
            Width           =   1515
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   14880
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   585
            Width           =   1515
         End
         Begin VB.TextBox Txtorder 
            Alignment       =   2  'Center
            Height          =   390
            Left            =   3390
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   150
            Width           =   1245
         End
         Begin VB.ComboBox CbBasedOn 
            Height          =   315
            ItemData        =   "FrmVocationEntitlements.frx":6BCC
            Left            =   5700
            List            =   "FrmVocationEntitlements.frx":6BCE
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   150
            Width           =   2730
         End
         Begin VB.TextBox TxtNoVacation 
            Alignment       =   2  'Center
            Height          =   390
            Left            =   540
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   150
            Width           =   1515
         End
         Begin VB.TextBox TxtStay 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1905
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox TxtCivilin 
            Alignment       =   2  'Center
            Height          =   390
            Left            =   405
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   150
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.ComboBox Contract_period 
            Height          =   315
            ItemData        =   "FrmVocationEntitlements.frx":6BD0
            Left            =   18540
            List            =   "FrmVocationEntitlements.frx":6BDA
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   0
            Visible         =   0   'False
            Width           =   1080
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   390
            Left            =   12390
            TabIndex        =   49
            Top             =   150
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   223674369
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Height          =   315
            Left            =   9705
            TabIndex        =   50
            Top             =   585
            Width           =   5070
            _ExtentX        =   8943
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmVocationEntitlements.frx":6BE8
            Height          =   315
            Left            =   5700
            TabIndex        =   51
            Top             =   585
            Width           =   2730
            _ExtentX        =   4815
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
         Begin MSDataListLib.DataCombo DcboJobsType 
            Height          =   315
            Left            =   9705
            TabIndex        =   52
            Top             =   1005
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker BignDate 
            Height          =   375
            Left            =   14880
            TabIndex        =   53
            Top             =   1440
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   223674369
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DateSta 
            Height          =   390
            Left            =   9705
            TabIndex        =   54
            Top             =   150
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   223674369
            CurrentDate     =   41640
         End
         Begin MSDataListLib.DataCombo Opretot 
            Bindings        =   "FrmVocationEntitlements.frx":6BFD
            Height          =   315
            Left            =   540
            TabIndex        =   55
            Top             =   585
            Width           =   4095
            _ExtentX        =   7223
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
         Begin MSComCtl2.DTPicker stratDate 
            Height          =   390
            Left            =   6915
            TabIndex        =   56
            Top             =   1005
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            Format          =   223674369
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker EndDate 
            Height          =   390
            Left            =   4320
            TabIndex        =   57
            Top             =   1005
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   223674369
            CurrentDate     =   41640
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   450
            Index           =   0
            Left            =   285
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1005
            Width           =   4050
            _cx             =   7144
            _cy             =   794
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
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   390
               Left            =   0
               TabIndex        =   251
               Top             =   0
               Visible         =   0   'False
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   688
               _Version        =   393216
               Format          =   223674369
               CurrentDate     =   41640
            End
            Begin XtremeSuiteControls.RadioButton Opt 
               Height          =   450
               Index           =   0
               Left            =   3090
               TabIndex        =   59
               Top             =   0
               Width           =   825
               _Version        =   786432
               _ExtentX        =   1455
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "≈Ã«“…"
               ForeColor       =   16711680
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Opt 
               Height          =   450
               Index           =   1
               Left            =   1470
               TabIndex        =   60
               Top             =   0
               Width           =   1515
               _Version        =   786432
               _ExtentX        =   2672
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "⁄·Ï —«” «·⁄„·"
               ForeColor       =   16711680
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Opt 
               Height          =   450
               Index           =   2
               Left            =   270
               TabIndex        =   61
               Top             =   0
               Width           =   1080
               _Version        =   786432
               _ExtentX        =   1905
               _ExtentY        =   794
               _StockProps     =   79
               Caption         =   "‰Â«Ì… Œœ„Â"
               ForeColor       =   16711680
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin MSDataListLib.DataCombo DcbDept 
            Height          =   315
            Left            =   9705
            TabIndex        =   62
            Top             =   1440
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker LastVocatinDate 
            Height          =   390
            Left            =   14880
            TabIndex        =   63
            Top             =   1005
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   223674369
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dctype 
            Height          =   315
            Left            =   4320
            TabIndex        =   64
            Top             =   1440
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   870
            Index           =   1
            Left            =   15810
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1875
            Width           =   2445
            _cx             =   4313
            _cy             =   1535
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
               Height          =   270
               Left            =   -420
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   0
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.TextBox TxtDayAbs 
               Alignment       =   2  'Center
               BackColor       =   &H00FF80FF&
               Enabled         =   0   'False
               Height          =   270
               Left            =   930
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   345
               Width           =   435
            End
            Begin VB.TextBox TxtYearAbs 
               Alignment       =   2  'Center
               BackColor       =   &H00FF80FF&
               Enabled         =   0   'False
               Height          =   270
               Left            =   105
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   345
               Width           =   435
            End
            Begin VB.TextBox TxtMoAbs 
               Alignment       =   2  'Center
               BackColor       =   &H00FF80FF&
               Enabled         =   0   'False
               Height          =   270
               Left            =   510
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   345
               Width           =   420
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   300
               Index           =   47
               Left            =   510
               TabIndex        =   73
               Top             =   105
               Width           =   345
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
               Height          =   300
               Index           =   46
               Left            =   105
               TabIndex        =   72
               Top             =   105
               Width           =   375
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÌÊ„"
               Height          =   300
               Index           =   45
               Left            =   1035
               TabIndex        =   71
               Top             =   105
               Width           =   255
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ì«„ «·€Ì«»"
               ForeColor       =   &H00C00000&
               Height          =   525
               Index           =   58
               Left            =   1560
               TabIndex        =   70
               Top             =   0
               Width           =   750
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   870
            Index           =   2
            Left            =   13110
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   1875
            Width           =   2685
            _cx             =   4736
            _cy             =   1535
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
               Height          =   270
               Left            =   105
               TabIndex        =   77
               Top             =   345
               Width           =   450
            End
            Begin VB.TextBox TxtMonth 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0C0&
               Enabled         =   0   'False
               Height          =   270
               Left            =   675
               TabIndex        =   76
               Top             =   345
               Width           =   510
            End
            Begin VB.TextBox TxtDay 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0C0&
               Enabled         =   0   'False
               Height          =   270
               Left            =   1215
               TabIndex        =   75
               Top             =   345
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   300
               Index           =   36
               Left            =   675
               TabIndex        =   81
               Top             =   105
               Width           =   285
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
               Height          =   300
               Index           =   28
               Left            =   105
               TabIndex        =   80
               Top             =   105
               Width           =   315
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÌÊ„"
               Height          =   300
               Index           =   35
               Left            =   1200
               TabIndex        =   79
               Top             =   105
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œ… «·Œœ„…"
               ForeColor       =   &H00C00000&
               Height          =   525
               Index           =   59
               Left            =   1785
               TabIndex        =   78
               Top             =   0
               Width           =   870
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   870
            Index           =   4
            Left            =   10530
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   1875
            Width           =   2565
            _cx             =   4524
            _cy             =   1535
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
               Height          =   330
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   -570
               Visible         =   0   'False
               Width           =   570
            End
            Begin VB.TextBox TxtVSa 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0FF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   1215
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   285
               Width           =   585
            End
            Begin VB.TextBox TxtYaerOut 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0FF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   135
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   285
               Width           =   555
            End
            Begin VB.TextBox TxtMontOut 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0FF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   675
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   285
               Width           =   555
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÌÊ„"
               Height          =   375
               Index           =   44
               Left            =   1365
               TabIndex        =   90
               Top             =   0
               Width           =   330
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   375
               Index           =   42
               Left            =   795
               TabIndex        =   89
               Top             =   0
               Width           =   330
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
               Height          =   375
               Index           =   43
               Left            =   135
               TabIndex        =   88
               Top             =   0
               Width           =   465
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã«“… »œÊ‰ —« »"
               ForeColor       =   &H00C00000&
               Height          =   795
               Index           =   60
               Left            =   2040
               TabIndex        =   87
               Top             =   0
               Width           =   465
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   870
            Index           =   5
            Left            =   8265
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   585
            Visible         =   0   'False
            Width           =   1740
            _cx             =   3069
            _cy             =   1535
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
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               Height          =   345
               Index           =   33
               Left            =   285
               TabIndex        =   93
               Top             =   420
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—« »"
               ForeColor       =   &H00C00000&
               Height          =   375
               Index           =   61
               Left            =   285
               TabIndex        =   92
               Top             =   0
               Visible         =   0   'False
               Width           =   1350
            End
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   7590
            TabIndex        =   144
            Top             =   0
            Visible         =   0   'False
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   223346689
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DataSalary 
            Height          =   375
            Left            =   0
            TabIndex        =   145
            Top             =   0
            Visible         =   0   'False
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            _Version        =   393216
            Format          =   223346689
            CurrentDate     =   41640
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   870
            Index           =   13
            Left            =   7740
            TabIndex        =   231
            TabStop         =   0   'False
            Top             =   1875
            Width           =   2775
            _cx             =   4895
            _cy             =   1535
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
            Begin VB.TextBox TxtDay2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0C0&
               Enabled         =   0   'False
               Height          =   330
               Left            =   1350
               TabIndex        =   234
               Top             =   240
               Width           =   480
            End
            Begin VB.TextBox TxtMonth2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0C0&
               Enabled         =   0   'False
               Height          =   330
               Left            =   720
               TabIndex        =   233
               Top             =   270
               Width           =   555
            End
            Begin VB.TextBox TxtYear2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0C0&
               Enabled         =   0   'False
               Height          =   330
               Left            =   105
               TabIndex        =   232
               Top             =   270
               Width           =   555
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œ… «·⁄„·"
               ForeColor       =   &H00C00000&
               Height          =   525
               Index           =   67
               Left            =   1920
               TabIndex        =   238
               Top             =   0
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÌÊ„"
               Height          =   300
               Index           =   66
               Left            =   1395
               TabIndex        =   237
               Top             =   -30
               Width           =   255
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
               Height          =   300
               Index           =   65
               Left            =   105
               TabIndex        =   236
               Top             =   -30
               Width           =   360
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   300
               Index           =   64
               Left            =   720
               TabIndex        =   235
               Top             =   -30
               Width           =   345
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   870
            Index           =   14
            Left            =   0
            TabIndex        =   243
            TabStop         =   0   'False
            Top             =   1875
            Width           =   2850
            _cx             =   5027
            _cy             =   1535
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
            Begin VB.TextBox TxtYear3 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0C0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   90
               TabIndex        =   246
               Top             =   285
               Width           =   480
            End
            Begin VB.TextBox TxtMonth3 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0C0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   645
               TabIndex        =   245
               Top             =   285
               Width           =   480
            End
            Begin VB.TextBox TxtDay3 
               Alignment       =   2  'Center
               BackColor       =   &H00FFC0C0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1185
               TabIndex        =   244
               Top             =   285
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   255
               Index           =   74
               Left            =   645
               TabIndex        =   250
               Top             =   90
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
               Height          =   255
               Index           =   73
               Left            =   90
               TabIndex        =   249
               Top             =   90
               Width           =   315
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÌÊ„"
               Height          =   255
               Index           =   72
               Left            =   1290
               TabIndex        =   248
               Top             =   90
               Width           =   210
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’«›Ì „œ… «·⁄„·"
               ForeColor       =   &H00C00000&
               Height          =   450
               Index           =   71
               Left            =   1740
               TabIndex        =   247
               Top             =   0
               Width           =   660
            End
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   375
            Left            =   0
            TabIndex        =   253
            Top             =   1440
            Visible         =   0   'False
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   661
            _Version        =   393216
            Format          =   223346689
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DTPicker5 
            Height          =   345
            Left            =   0
            TabIndex        =   254
            Top             =   585
            Visible         =   0   'False
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   609
            _Version        =   393216
            Format          =   223346689
            CurrentDate     =   41640
         End
         Begin XtremeSuiteControls.CheckBox Ch 
            Height          =   330
            Index           =   9
            Left            =   5430
            TabIndex        =   267
            Top             =   2280
            Width           =   285
            _Version        =   786432
            _ExtentX        =   503
            _ExtentY        =   582
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ì«„ «· Ï  ’—›"
            ForeColor       =   &H00000000&
            Height          =   555
            Index           =   78
            Left            =   6510
            TabIndex        =   266
            Top             =   2220
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—’Ìœ „—Õ·"
            ForeColor       =   &H00000000&
            Height          =   555
            Index           =   77
            Left            =   0
            TabIndex        =   265
            Top             =   0
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»œÊ‰ —« »"
            Height          =   405
            Index           =   76
            Left            =   11445
            TabIndex        =   263
            Top             =   2850
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ì«„ „Œ’Ê„Â"
            ForeColor       =   &H00000000&
            Height          =   555
            Index           =   70
            Left            =   4065
            TabIndex        =   242
            Top             =   1875
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘Â—"
            ForeColor       =   &H00C00000&
            Height          =   555
            Index           =   69
            Left            =   5160
            TabIndex        =   240
            Top             =   1875
            Width           =   510
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—’Ìœ „—Õ·"
            ForeColor       =   &H00000000&
            Height          =   555
            Index           =   68
            Left            =   6495
            TabIndex        =   239
            Top             =   1875
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ì«„ «·«Ã«“… «·„” Õﬁ… ﬁ»· «·Œ’„"
            Height          =   540
            Index           =   2
            Left            =   16260
            TabIndex        =   117
            Top             =   2865
            Width           =   1995
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ì«„ «·€Ì«»"
            Height          =   405
            Index           =   48
            Left            =   14700
            TabIndex        =   116
            Top             =   2865
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»œÊ‰ —« »"
            Height          =   405
            Index           =   50
            Left            =   13110
            TabIndex        =   115
            Top             =   2865
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·«Ì«„ "
            Height          =   405
            Index           =   22
            Left            =   5190
            TabIndex        =   114
            Top             =   2865
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   330
            Index           =   31
            Left            =   3255
            TabIndex        =   113
            Top             =   2730
            Width           =   705
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ì«„ «·«Ã«“… «·„” Õﬁ…"
            Height          =   405
            Index           =   49
            Left            =   7080
            TabIndex        =   112
            Top             =   2865
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "’«›Ì «·€Ì«»"
            Height          =   405
            Index           =   51
            Left            =   9510
            TabIndex        =   111
            Top             =   2865
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—’Ìœ ≈Ã«“… ”«»ﬁ"
            Height          =   345
            Index           =   21
            Left            =   18810
            TabIndex        =   110
            Top             =   0
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—…/«·ﬁ”„"
            Height          =   330
            Index           =   0
            Left            =   13785
            TabIndex        =   109
            Top             =   1440
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ⁄ÌÌ‰"
            Height          =   330
            Index           =   9
            Left            =   16935
            TabIndex        =   108
            Top             =   1440
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ ‰Â«Ì… «·Œœ„Â"
            Height          =   330
            Index           =   55
            Left            =   8415
            TabIndex        =   107
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   360
            Index           =   1
            Left            =   13560
            TabIndex        =   106
            Top             =   150
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   315
            Index           =   3
            Left            =   16935
            TabIndex        =   105
            Top             =   585
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   360
            Index           =   4
            Left            =   16935
            TabIndex        =   104
            Top             =   180
            Width           =   1380
         End
         Begin VB.Label lblbr 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   315
            Left            =   8415
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   585
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÊŸÌ›…"
            Height          =   315
            Index           =   24
            Left            =   13785
            TabIndex        =   102
            Top             =   1005
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «Œ— «Ã«“…"
            Height          =   360
            Index           =   20
            Left            =   16935
            TabIndex        =   101
            Top             =   1005
            Width           =   1380
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "  «· ”ÊÌ…"
            Height          =   330
            Index           =   0
            Left            =   11070
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   150
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬁ«∆„"
            Height          =   285
            Index           =   23
            Left            =   4740
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   585
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»œ«Ì… «·«Ã«“…"
            Height          =   360
            Left            =   8415
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   1005
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Â«Ì…«·«Ã«“…"
            Height          =   315
            Left            =   5565
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   1005
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰« ⁄·Ï"
            Height          =   360
            Index           =   56
            Left            =   8415
            TabIndex        =   96
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ÿ·»"
            Height          =   330
            Left            =   4740
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «Ì«„ «·«Ã«“…"
            Height          =   330
            Left            =   2025
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   150
            Width           =   1245
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1215
         Index           =   11
         Left            =   6720
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   8865
         Width           =   11655
         _cx             =   20558
         _cy             =   2143
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
         Begin VB.CommandButton Command2 
            Caption         =   "Õ–› ﬁÌœ «·«” Õﬁ«ﬁ"
            Height          =   390
            Left            =   7635
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   300
            Width           =   1890
         End
         Begin VB.CommandButton Command5 
            Caption         =   "≈‰‘«¡ ﬁÌœ «·«” Õﬁ«ﬁ"
            Height          =   390
            Left            =   9645
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   300
            Width           =   1950
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   13560
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   -165
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   390
            Left            =   3810
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   300
            Width           =   2595
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   390
            Left            =   1905
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   300
            Width           =   1395
         End
         Begin VB.CommandButton Command8 
            Caption         =   "ﬂ‘› Õ”«»"
            Height          =   390
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   285
            Width           =   1515
         End
         Begin XtremeSuiteControls.CheckBox chkGE 
            Height          =   660
            Left            =   9615
            TabIndex        =   143
            Top             =   600
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   1164
            _StockProps     =   79
            Caption         =   "«·ﬁÌœ ⁄·Ì  «—ÌŒ «·Õ—ﬂ…"
            ForeColor       =   8388608
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo ADDACC 
            Bindings        =   "FrmVocationEntitlements.frx":6C12
            Height          =   315
            Left            =   5025
            TabIndex        =   255
            Top             =   795
            Width           =   3150
            _ExtentX        =   5556
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
         Begin MSDataListLib.DataCombo DISACC 
            Bindings        =   "FrmVocationEntitlements.frx":6C27
            Height          =   315
            Left            =   135
            TabIndex        =   257
            Top             =   690
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õ Œ’Ê„«  «Œ—Ì"
            Height          =   480
            Index           =   2
            Left            =   3525
            RightToLeft     =   -1  'True
            TabIndex        =   258
            Top             =   735
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õ «÷«›«  «Œ—Ì"
            Height          =   465
            Index           =   1
            Left            =   8250
            RightToLeft     =   -1  'True
            TabIndex        =   256
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   435
            Index           =   35
            Left            =   6225
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   300
            Width           =   1290
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   285
         TabIndex        =   130
         Top             =   9615
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   660
         Left            =   0
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   10080
         Width           =   18420
         _cx             =   32491
         _cy             =   1164
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
            Height          =   420
            Index           =   0
            Left            =   16995
            TabIndex        =   133
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   741
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
            Height          =   420
            Index           =   1
            Left            =   15810
            TabIndex        =   134
            Top             =   120
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   741
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
            Height          =   420
            Index           =   2
            Left            =   14535
            TabIndex        =   135
            Top             =   120
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   741
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
            Height          =   420
            Index           =   3
            Left            =   13305
            TabIndex        =   136
            Top             =   120
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   741
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
            Height          =   420
            Index           =   4
            Left            =   11790
            TabIndex        =   137
            Top             =   120
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   741
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
            Height          =   420
            Index           =   6
            Left            =   150
            TabIndex        =   138
            Top             =   120
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   741
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
            Height          =   420
            Left            =   3720
            TabIndex        =   139
            Top             =   120
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   741
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
            Height          =   420
            Index           =   5
            Left            =   10380
            TabIndex        =   140
            Top             =   120
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   741
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
            Height          =   420
            Index           =   9
            Left            =   7980
            TabIndex        =   141
            Top             =   120
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   741
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «·„” Õﬁ« "
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
            ButtonImage     =   "FrmVocationEntitlements.frx":6C3C
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
            Height          =   420
            Index           =   8
            Left            =   5910
            TabIndex        =   142
            Top             =   120
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   741
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «· ⁄Âœ"
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
            ButtonImage     =   "FrmVocationEntitlements.frx":D49E
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Accredit 
            Height          =   420
            Left            =   1800
            TabIndex        =   224
            Top             =   120
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   741
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   420
            Index           =   11
            Left            =   4680
            TabIndex        =   261
            Top             =   120
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   741
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ… : "
         Height          =   300
         Index           =   8
         Left            =   3435
         TabIndex        =   131
         Top             =   9705
         Width           =   1005
      End
      Begin VB.Label XPTxtCurrent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   390
         Left            =   4515
         TabIndex        =   129
         Top             =   9315
         Width           =   720
      End
      Begin VB.Label XPTxtCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   390
         Left            =   2715
         TabIndex        =   128
         Top             =   9315
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ⁄œœ «·”Ã·« :"
         Height          =   390
         Index           =   6
         Left            =   3360
         TabIndex        =   127
         Top             =   9315
         Width           =   1080
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·”Ã· «·Õ«·Ì:"
         Height          =   390
         Index           =   7
         Left            =   5280
         TabIndex        =   126
         Top             =   9315
         Width           =   1200
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1650
      Visible         =   0   'False
      Width           =   975
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
      TabIndex        =   0
      Top             =   4770
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   120
      Top             =   4680
      Width           =   6255
   End
End
Attribute VB_Name = "FrmVocationEntitlements"
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
Public bol As Boolean
Public novalue As Boolean
Dim cProgress As ClsProgress
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim FixedOrChanged(40) As Integer
Dim AddOrDiscount(40) As Integer
Dim ViewComp(40) As Boolean
Dim showMofradAll(40) As Boolean
Dim culc30orRminder(40) As Integer
Dim Account_code(40) As String
Dim Account_code1(40) As String
Dim ZmamAccount(40) As String
Dim AdvPaymentdAccount(40) As String
Dim componentname(40) As String
Dim firstrun As Boolean
'Private Sub Accredit_Click()
'    Dim BeginTrans As Boolean
'
'    Cn.BeginTrans
'    BeginTrans = True
'
'    If IsNull(rs("Posted")) Then
'        rs("Posted") = user_id
'        rs("PostedDate") = Time
'    Else
'        rs("Posted") = Null
'       rs("PostedDate") = Time
'    End If
'
'    rs.update
' If SystemOptions.UserInterface = ArabicInterface Then
'    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
'Else
'Accredit.Caption = "Sent To approval "
'End If
'
'    Cn.CommitTrans
''    BeginTrans = False
'FillApprovedTable
'    Retrive (val(Me.XPTxtID.text))
'End Sub
Sub GetAbsenece(Optional EmID As Integer, Optional Typ As Integer, Optional ByRef Total As Double = 0)
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = " SELECT     SUM(NoDay) AS Sm"
sql = sql & " From dbo.TblInforVacatiom"
sql = sql & " where (EmpID=" & EmID & ") and (TypeVacation=" & Typ & ")"
Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs7.RecordCount >= 0 Then
Total = IIf(IsNull(Rs7("Sm").value), 0, Rs7("Sm").value)
End If
End Sub
Sub SaveInformationVacation(Optional TypeVacation As Integer = 0, Optional EmpID As Integer = 0, Optional NoDay As Double = 0)
Dim sql As String
Dim str As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If SystemOptions.UserInterface = ArabicInterface Then
str = " „” Õﬁ«  «·ﬁ«Ì„ [Ã«“…"
Else
str = "Due to Vacation"
End If
sql = "select * from TblInforVacatiom where (1=-1)"
    Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      Rs7.AddNew
      Rs7("VacatioID").value = val(XPTxtID.text)
      Rs7("EmpID").value = EmpID
      Rs7("NoDay").value = (NoDay * -1)
      Rs7("RecordDate").value = XPDtbTrans.value
      Rs7("RecordDateH").value = ToHijriDate(XPDtbTrans.value)
      Rs7("TypeVacation").value = TypeVacation
      Rs7("Remarks").value = str
      Rs7.update
End Sub

Sub ChekVacation()
Dim PeriodMonth As Double
Dim Tempperiod As Double
Dim Period As Double
Dim HoldyNo As Double
Dim NODiffDate As Double
Dim TempValu As Double
    If CheckSettingsVacType() = True Then
    
    GetHoldayDays2 val(DcboEmpName.BoundText), Period, HoldyNo
   ' PeriodMonth = DateDiff("d", LastVocatinDate.value, DateSta.value)
   PeriodMonth = val(TxtYear2.text) * 12 * 30 + val(TxtMonth2.text) * 30 + val(TxtDay2.text)
    Tempperiod = Period
    
    
    If Period <> 0 Then
    If CheckSettingsLikeContract() = True Then
    NODiffDate = PeriodMonth + (GetLastBalanceMonthVaction(val(DcboEmpName.BoundText), val(XPTxtID.text)) * 30) - val(TxtDiscouDay.text)
    
    Period = Period * 30
   If NODiffDate >= Period Then
   TempValu = NODiffDate \ Period
    TxtContDay.text = Round((HoldyNo / Period) * NODiffDate, 2) 'HoldyNo * TempValu
    LastBalanceMonth.text = Round((NODiffDate - Period * TempValu) / 30, 2)
   Else
   LastBalanceMonth.text = 0
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
StrSQL = StrSQL & " WHERE     (EmpID = " & val(Me.DcboEmpName.BoundText) & ") AND (ExpectedacationDate <=" & SQLDate(DateSta.value, True) & ") AND (Status1 IS NULL)"
StrSQL = StrSQL & " GROUP BY EmpID"
  Rs3.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
  If Rs3.RecordCount > 0 Then
 TxtContDay.text = IIf(IsNull(Rs3("Tota").value), 0, Rs3("Tota").value)
 End If
  End If
End Sub
Function GetHobStatus() As Integer
Dim sql As String
GetHobStatus = 0
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     id, Vacation"
sql = sql & " From dbo.jopstatus"
sql = sql & " Where (Vacation = 1)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetHobStatus = IIf(IsNull(Rs3("id").value), 0, Rs3("id").value)
Else
GetHobStatus = 0
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
 
    SendTopost Me.Name, "TblVocationEntitlements", "Id", val(DcbDept.BoundText), val(Dcbranch.BoundText), val(XPTxtID.text), XPTxtID
   rs.Resync
   
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
    Retrive (val(Me.XPTxtID.text))
End Sub

Private Sub ADDACC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 6660
    End If

End Sub

Private Sub Ch_Click(index As Integer)
Smation
End Sub

Private Sub ChDelivery_Click()
If ChDelivery.value = vbChecked Then
Total.text = val(NetTotal.text) + val(TxtValueTickt.text)
Else
Total.text = val(NetTotal.text)
End If
End Sub


Private Sub Cmd_Click(index As Integer)

    ' On Error GoTo ErrTrap
    Select Case index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            DcboEmpName.BoundText = ""
            lbl(10).Caption = "0"
            clear_all Me
              VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
              VSFlexGrid2.rows = 1
              Grid.Clear flexClearScrollable, flexClearEverything
              Grid.rows = 1
              Grid1.Clear flexClearScrollable, flexClearEverything
              Grid1.rows = 1
              
            Accredit.Caption = ""
            Accredit.Enabled = False
            
            Me.DCboUserName.BoundText = user_id
            Opretot.BoundText = user_id
        '    TxtPaymentCounts.text = 1
        Dcbranch.BoundText = Current_branch
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
           VSFlexGrid1.rows = 1
           Option1.value = True
           Opt(0).value = True
            'XPDtbTrans.SetFocus
            
'            Accredit.Enabled = True
'                If SystemOptions.UserInterface = ArabicInterface Then
'                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
'                                                  Else
'                                                    Accredit.Caption = " send to Approval   "
'                                               End If
             
             
        Case 1
        If TxtNoteSerial.text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· .Ì—ÃÏ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ"
        Else
        MsgBox "Can Not edit .Delete JE"
        End If
        Exit Sub
        End If
        If chkGE.value = xtpChecked Then
          If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
        Else
              
          If ChekClodePeriod(DateSta.value) = True Then
                         If SystemOptions.UserInterface = ArabicInterface Then
                          MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «· ”ÊÌ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                         Else
                         MsgBox "Please Change Date Becouse This is Period is Closed"
                        End If
              Exit Sub
         End If
     End If
     
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
             If CheWork() = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·Õ—ﬂ… „— »ÿ… »„»«‘—… ⁄„·"
            Else
            MsgBox "Can not edite this process Linked to the initiation of work"
            End If
            Exit Sub
            End If
            
    If CheWork() = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·Õ—ﬂ… „— »ÿ… »„»«‘—… ⁄„·"
            Else
            MsgBox "Can not edite this process Linked to the initiation of work"
            End If
            Exit Sub
            End If
            
           If ChePayment() = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· Â–Â «·Õ—ﬂ… „— »ÿ… »«·„œ›Ê⁄« "
            Else
            MsgBox "Can not edite this process Linked to Payments"
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
            Opretot.BoundText = user_id

        Case 2

    
         If chkGE.value = xtpChecked Then
          If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
        Else
              
          If ChekClodePeriod(DateSta.value) = True Then
                         If SystemOptions.UserInterface = ArabicInterface Then
                          MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «· ”ÊÌ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                         Else
                         MsgBox "Please Change Date Becouse This is Period is Closed"
                        End If
              Exit Sub
         End If
     End If
     
              
            Dim Msg As String

            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Dcbranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.Dcbranch.BoundText
If GetHobStatus() = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "Ì—ÃÏ  ÕœÌœ «·«Ã«“… „‰ ‘«‘… Õ«·«  «·⁄„·"
Else
Msg = "Please Select Vacaion From Screen Job Situations"
End If
MsgBox Msg
Exit Sub
End If
''/////////////////   If CheckSettingsVacType() = True Then
Dim idd As Double
 If CheckSettingsVacType() = True Then
    Dim Period As Double
    Dim Diff As Double
    
    ''/////////
    
    Diff = DateDiff("M", LastVocatinDate.value, DateSta.value)
   
    If CheckSettingsLikeContract() = True Then
    GetHoldayDays2 val(DcboEmpName.BoundText), Period
    Diff = val(TxtYear2.text) * 12 * 30 + val(TxtMonth2.text) * 30 + val(TxtDay2.text)
  '  Diff = DateDiff("d", LastVocatinDate.value, DateSta.value)
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
     DTPicker1.value = DateAdd("M", Period, LastVocatinDate.value)
     If GetSettingsVacDate(DTPicker1.value, idd) = True Then
     Else
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«· «—ÌŒ €Ì— „ÿ«»ﬁ ·«⁄œ«œ«  «·«Ã«“« "
     Else
     MsgBox "Date does not match vacation settings"
     End If
     Exit Sub
     End If

     If GetSettingsVacDateAllow(DateSta.value, idd) = True Then
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
If CHeckPayedSalaryCurrMonth() = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "ÌÊÃœ ﬁÌœ «” Õﬁ«ﬁ ·Â–« «·‘Â— Â·  —Ìœ «·„Ê«’·… "
Else
MsgBox "There is an entitlement for this month"
End If
If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
Else
Exit Sub
End If
End If

            SaveData

        Case 3
            Undo

        Case 4
             If TxtNoteSerial.text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «·Õ–› .Ì—ÃÏ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ"
        Else
        MsgBox "Can Not delete .Delete JE"
        End If
        Exit Sub
        End If
 
          If chkGE.value = xtpChecked Then
          If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
        Else
              
          If ChekClodePeriod(DateSta.value) = True Then
                         If SystemOptions.UserInterface = ArabicInterface Then
                          MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «· ”ÊÌ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                         Else
                         MsgBox "Please Change Date Becouse This is Period is Closed"
                        End If
              Exit Sub
         End If
     End If
      If ScreenAproved(val(XPTxtID.text), Me.Name) = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
         MsgBox "·«Ì„ﬂ‰ «·Õ–›.Â–Â «·Õ—ﬂ… „— »ÿ… »«·«⁄ „«œ« "
         Else
         MsgBox "Can not delete.This process associated with approvals"
         End If
         Exit Sub
       End If
       
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
        bol = True

        FrmSearchVocationEntitlement.index = 0
            Load FrmSearchVocationEntitlement
          FrmSearchVocationEntitlement.show vbModal
            

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 8

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
               print_report val(Me.XPTxtID.text), 1
            End If
            
       Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text), 0
            End If
      Case 11
                  
                        On Error Resume Next
ShowAttachments XPTxtID.text, "2312202002"

    End Select

    Exit Sub
ErrTrap:
End Sub
'Function print_reportt(Optional NoteSerial As String)
'
'    Dim MySQL As String
'    Dim RsData As New ADODB.Recordset
'    Dim xApp As New CRAXDRT.Application
'    Dim xReport As CRAXDRT.Report
'    Dim CViewer As ClsReportViewer
'    Dim StrReportTitle As String
'    Dim StrFileName As String
'    Dim Msg As String
'
'
'
'MySQL = "SELECT     dbo.TblVocationEntitlements.RecordDate, dbo.TblVocationEntitlements.DateSta, dbo.TblVocationEntitlements.OpretotID, dbo.TblUsers.UserName, "
'MySQL = MySQL & "                      dbo.TblVocationEntitlements.EmpID, TblEmployee_2.Emp_Name, TblEmployee_2.Emp_Name1, TblEmployee_2.Emp_Name2, TblEmployee_2.Emp_Name3,"
'MySQL = MySQL & "                      TblEmployee_2.Emp_Name4, TblEmployee_2.Nationality, TblEmployee_2.Fullcode, TblEmployee_2.Emp_Namee4, TblEmployee_2.Emp_Namee3,"
'MySQL = MySQL & "                      TblEmployee_2.Emp_Namee2, TblEmployee_2.Emp_Namee1, TblEmployee_2.Emp_Namee, dbo.TblVocationEntitlements.BranchID,"
'''MySQL = MySQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblVocationEntitlements.JobID, dbo.TblEmpJobsTypes.JobTypeName,"
'm 'ySQL = MySQL & "                      dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblVocationEntitlements.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
'MySQL = MySQL & "                      dbo.TblVocationEntitlements.BignDate, dbo.TblVocationEntitlements.LastVocatinDate, dbo.TblVocationEntitlements.ContDay, dbo.TblVocationEntitlements.LastDayVoc,"
'MySQL = MySQL & "                      dbo.TblVocationEntitlements.TotalDay, dbo.TblVocationEntitlements.NoDay, dbo.TblVocationEntitlements.NoMonth, dbo.TblVocationEntitlements.NoYear,"
'MySQL = MySQL & "                      dbo.TblVocationEntitlements.Remark, dbo.TblVocationEntitlements.DaySalary, dbo.TblVocationEntitlements.Salary, dbo.TblVocationEntitlements.DayIncrease,"
'MySQL = MySQL & "                      dbo.TblVocationEntitlements.Increase, dbo.TblVocationEntitlements.DaySalVocation, dbo.TblVocationEntitlements.SalaryVocation,"
'MySQL = MySQL & "                      dbo.TblVocationEntitlements.DayEntitOther, dbo.TblVocationEntitlements.SalEntitOther, dbo.TblVocationEntitlements.Other, dbo.TblVocationEntitlements.Advance,"
'MySQL = MySQL & "                      dbo.TblVocationEntitlements.ValueTickt, dbo.TblVocationEntitlements.Booked, dbo.TblVocationEntitlements.Delivery, dbo.TblVocationEntitlements.ID,"
'MySQL = MySQL & "                      dbo.TblVocationEntitlementsDet.DeliverDate, dbo.TblVocationEntitlementsDet.ReciveDate, dbo.TblVocationEntitlementsDet.Valu,"
'MySQL = MySQL & "                      dbo.TblVocationEntitlementsDet.TypeM, dbo.TblVocationEntitlementsDet.MofrdID, dbo.mofrad.name, dbo.mofrad.nameE,"
'MySQL = MySQL & "                      dbo.TblVocationEntitlementsDet.EmpID AS EmpIDDet, TblEmployee_1.Emp_Name AS Emp_NameDet, TblEmployee_1.Emp_Name1 AS Emp_NameDet1,"
'MySQL = MySQL & "                      TblEmployee_1.Emp_Name2 AS Emp_NameDet2, TblEmployee_1.Emp_Name3 AS Emp_NameDet3, TblEmployee_1.Emp_Name4 AS Emp_NameDet4,"
'MySQL = MySQL & "                      TblEmployee_1.Fullcode AS FullcodeDet, TblEmployee_1.Emp_Namee4 AS Emp_Namee4Det4, TblEmployee_1.Emp_Namee3 AS Emp_Namee4Det3,"
'MySQL = MySQL & "                      TblEmployee_1.Emp_Namee2 AS Emp_Namee4Det2, TblEmployee_1.Emp_Namee1 AS Emp_Namee4Det1, TblEmployee_1.Emp_Namee AS Emp_Namee4Det,"
'MySQL = MySQL & "                      dbo.TblAssestes.AsName, dbo.TblAssestes.AsCode, dbo.TblVocationEntitlements.DuVocation, dbo.TblVocationEntitlements.ToalAbsent,"
'MySQL = MySQL & "                      dbo.TblVocationEntitlements.YearAbs, dbo.TblVocationEntitlements.MoAbs, dbo.TblVocationEntitlements.DayAbs, dbo.TblVocationEntitlements.DayOut,"
'MySQL = MySQL & "                      dbo.TblVocationEntitlements.MontOut , dbo.TblVocationEntitlements.YaerOut"
'MySQL = MySQL & " FROM         dbo.TblEmpDepartments RIGHT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblAssestes RIGHT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblVocationEntitlementsDet ON dbo.TblAssestes.AsID = dbo.TblVocationEntitlementsDet.MofrdID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblVocationEntitlementsDet.EmpID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.mofrad ON dbo.TblVocationEntitlementsDet.MofrdID = dbo.mofrad.id RIGHT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblVocationEntitlements ON dbo.TblVocationEntitlementsDet.VoEntID = dbo.TblVocationEntitlements.ID ON"
'MySQL = MySQL & "                      dbo.TblEmpDepartments.DeparmentID = dbo.TblVocationEntitlements.DeptID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TblVocationEntitlements.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblVocationEntitlements.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.TblVocationEntitlements.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
'MySQL = MySQL & "                      dbo.TblUsers ON dbo.TblVocationEntitlements.OpretotID = dbo.TblUsers.UserID"
'
'MySQL = MySQL & " Where (dbo.TblVocationEntitlements.id = " & val(XPTxtID.text) & ")"
'
'        If SystemOptions.UserInterface = ArabicInterface Then
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepVocationEnitylement2.rpt"
'        Else
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepVocationEnitylement2.rpt"
'        End If
'
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
'        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
'        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
'        StrReportTitle = ""
'        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
'        'End If
'        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
'        'End If
'    End If
'
'    xReport.ParameterFields(3).AddCurrentValue user_name
'      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'     '   xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
'        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    'xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
'' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
' ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
'  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
'
''    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
'    xReport.reporttitle = StrReportTitle
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.Title
'    xReport.ReportAuthor = App.Title
'    Set CViewer = New ClsReportViewer
'    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
'
'    RsData.Close
'    Set RsData = Nothing
'    Screen.MousePointer = vbDefault
'
'
'
'
'
'End Function
Function print_report(Optional NoteSerial As String, Optional index As Integer = 0)
        
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT   TblEmployee_2.NumEkama,  dbo.TblVocationEntitlements.RecordDate, dbo.TblVocationEntitlements.DateSta, dbo.TblVocationEntitlements.OpretotID, dbo.TblUsers.UserName, "
MySQL = MySQL & "                      dbo.TblVocationEntitlements.EmpID, TblEmployee_2.Emp_Name, TblEmployee_2.Emp_Name1, TblEmployee_2.Emp_Name2, TblEmployee_2.Emp_Name3,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Name4, TblEmployee_2.Nationality, TblEmployee_2.Fullcode, TblEmployee_2.Emp_Namee4, TblEmployee_2.Emp_Namee3,"
MySQL = MySQL & "                      TblEmployee_2.Emp_Namee2, TblEmployee_2.Emp_Namee1, TblEmployee_2.Emp_Namee, dbo.TblVocationEntitlements.BranchID,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblVocationEntitlements.JobID, dbo.TblEmpJobsTypes.JobTypeName,"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblVocationEntitlements.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.BignDate, dbo.TblVocationEntitlements.LastVocatinDate, dbo.TblVocationEntitlements.ContDay, dbo.TblVocationEntitlements.LastDayVoc,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.TotalDay, dbo.TblVocationEntitlements.NoDay, dbo.TblVocationEntitlements.NoMonth, dbo.TblVocationEntitlements.NoYear,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.Remark, dbo.TblVocationEntitlements.DaySalary, dbo.TblVocationEntitlements.Salary, dbo.TblVocationEntitlements.DayIncrease,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.Increase, dbo.TblVocationEntitlements.DaySalVocation, dbo.TblVocationEntitlements.SalaryVocation,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.DayEntitOther, dbo.TblVocationEntitlements.SalEntitOther, dbo.TblVocationEntitlements.Other, dbo.TblVocationEntitlements.Advance,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.ValueTickt, dbo.TblVocationEntitlements.Booked, dbo.TblVocationEntitlements.Delivery, dbo.TblVocationEntitlements.ID,"
MySQL = MySQL & "                      dbo.TblVocationEntitlementsDet.DeliverDate, dbo.TblVocationEntitlementsDet.ReciveDate, dbo.TblVocationEntitlementsDet.Valu,"
MySQL = MySQL & "                      dbo.TblVocationEntitlementsDet.TypeM, dbo.TblVocationEntitlementsDet.MofrdID, dbo.mofrad.name, dbo.mofrad.nameE,"
MySQL = MySQL & "                      dbo.TblVocationEntitlementsDet.EmpID AS EmpIDDet, TblEmployee_1.Emp_Name AS Emp_NameDet, TblEmployee_1.Emp_Name1 AS Emp_NameDet1,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Name2 AS Emp_NameDet2, TblEmployee_1.Emp_Name3 AS Emp_NameDet3, TblEmployee_1.Emp_Name4 AS Emp_NameDet4,"
MySQL = MySQL & "                      TblEmployee_1.Fullcode AS FullcodeDet, TblEmployee_1.Emp_Namee4 AS Emp_Namee4Det4, TblEmployee_1.Emp_Namee3 AS Emp_Namee4Det3,"
MySQL = MySQL & "                      TblEmployee_1.Emp_Namee2 AS Emp_Namee4Det2, TblEmployee_1.Emp_Namee1 AS Emp_Namee4Det1, TblEmployee_1.Emp_Namee AS Emp_Namee4Det,"
MySQL = MySQL & "                      dbo.TblAssestes.AsName, dbo.TblAssestes.AsCode, dbo.TblVocationEntitlements.DuVocation, dbo.TblVocationEntitlements.ToalAbsent,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.YearAbs, dbo.TblVocationEntitlements.MoAbs, dbo.TblVocationEntitlements.DayAbs, dbo.TblVocationEntitlements.DayOut,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.MontOut, dbo.TblVocationEntitlements.YaerOut, dbo.TblVocationEntitlements.Chekk, dbo.TblVocationEntitlements.stratDate,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.EndDate, dbo.TblVocationEntitlements.ch5, dbo.TblVocationEntitlements.ch4, dbo.TblVocationEntitlements.ch3,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.ch2, dbo.TblVocationEntitlements.ch1, dbo.TblVocationEntitlements.ch0, dbo.TblVocationEntitlements.InsuranceValue,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.ch7, TblEmployee_2.BankCard, TblEmployee_2.BankCode, TblEmployee_2.BankIAddress, TblEmployee_2.BanckName,"
MySQL = MySQL & "                      TblEmployee_2.BankIBan, TblEmployee_2.DOB, dbo.TblVocationEntitlements.ch8, dbo.TblVocationEntitlements.PreSalary, dbo.TblVocationEntitlements.chkGE,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.PayedPayment, dbo.TblVocationEntitlements.TotalDue, dbo.TblVocationEntitlements.NetDue, dbo.TblVocationEntitlements.TotalCut,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.NetTotal, dbo.TblVocationEntitlements.decrease, dbo.TblVocationEntitlements.NoteSerial, dbo.TblVocationEntitlements.Vact_Work,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.TypEndService, dbo.TblVocationEntitlements.GetInsurance, dbo.TblVocationEntitlements.LastBalanceMonth,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.TxtNoVaction, dbo.TblVocationEntitlements.IDes, dbo.TblVocationEntitlements.TxtDay2, dbo.TblVocationEntitlements.TxtMonth2,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.TxtYear2, dbo.TblVocationEntitlements.TxtDay3, dbo.TblVocationEntitlements.TxtMonth3, dbo.TblVocationEntitlements.TxtYear3,"
MySQL = MySQL & "                      dbo.TblVocationEntitlements.TxtAddDay , dbo.TblVocationEntitlements.TxtDiscouDay, dbo.TblVocationEntitlements.ch6"
MySQL = MySQL & " FROM         dbo.TblEmpDepartments RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblAssestes RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblVocationEntitlementsDet ON dbo.TblAssestes.AsID = dbo.TblVocationEntitlementsDet.MofrdID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblVocationEntitlementsDet.EmpID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.mofrad ON dbo.TblVocationEntitlementsDet.MofrdID = dbo.mofrad.id RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblVocationEntitlements ON dbo.TblVocationEntitlementsDet.VoEntID = dbo.TblVocationEntitlements.ID ON"
MySQL = MySQL & "                      dbo.TblEmpDepartments.DeparmentID = dbo.TblVocationEntitlements.DeptID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpJobsTypes ON dbo.TblVocationEntitlements.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblVocationEntitlements.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.TblVocationEntitlements.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblUsers ON dbo.TblVocationEntitlements.OpretotID = dbo.TblUsers.UserID"
MySQL = MySQL & " Where (dbo.TblVocationEntitlements.id = " & val(XPTxtID.text) & ")"
If index = 0 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepVocationEnitylement.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepVocationEnitylementE.rpt"
        End If
   Else
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepVocationEnitylement2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepVocationEnitylement2.rpt"
        End If
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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
     '   xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    'xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
If index = 0 Then
 Dim SalaryStr As String
 Dim j As Integer
 Dim ColumnName As String
 Dim SngTotal As Double
 SalaryStr = ""
 With Grid
        For j = 1 To 40
        SngTotal = 0
            ColumnName = "Comp" & j
            SngTotal = val(.TextMatrix(.rows - 1, .ColIndex(ColumnName)))
            If SngTotal <> 0 Then
           SalaryStr = SalaryStr & .TextMatrix(0, .ColIndex(ColumnName)) & " =  " & SngTotal
           SalaryStr = SalaryStr & CHR(13)
           End If
        Next j
        xReport.ParameterFields(5).AddCurrentValue SalaryStr
      
  End With
  End If
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

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub





Private Sub Command1_Click()
Ele(10).Visible = True
C1Elastic2.Visible = False
Ele(12).Visible = False
ShowComponent
End Sub

Private Sub Command2_Click()
Dim Msg As String
Dim StrSQL As String
Dim X As Integer
         If chkGE.value = xtpChecked Then
          If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
        Else
              
          If ChekClodePeriod(DateSta.value) = True Then
                         If SystemOptions.UserInterface = ArabicInterface Then
                          MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «· ”ÊÌ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                         Else
                         MsgBox "Please Change Date Becouse This is Period is Closed"
                        End If
              Exit Sub
         End If
     End If
            If ChePayment() = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… „— »ÿ… »«·„œ›Ê⁄« "
            Else
            MsgBox "Can not Delete this process Linked to Payments"
            End If
            Exit Sub
            End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
    Else
        Msg = "Confirm Delete  "
    End If
        X = MsgBox(Msg, vbCritical + vbYesNo)

      If X = vbYes Then
        StrSQL = "Update TblEmployee Set  jopstatusid=1,workstate=0 Where Emp_ID=" & val(DcboEmpName.BoundText) & ""
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "update   TblVocationEntitlements set NoteSerial=null,NoteID=null Where ID=" & val(Me.XPTxtID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        rs.Resync
        Retrive (val(Me.XPTxtID.text))
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› ﬁÌœ «·«” Õﬁ«ﬁ  "
        Else
            Msg = " This voucher deleted  "
        End If
        MsgBox Msg
       End If
        
End Sub

Private Sub Command3_Click()
Ele(10).Visible = False
C1Elastic2.Visible = False
Ele(12).Visible = True
'FillGridWithData3
If val(val(DcboEmpName.BoundText)) = 0 Or DcboEmpName.text = "" Then Exit Sub
If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
ClaCulte
Else
FillGridSalarDat
End If

End Sub
Sub SaveSalary()
Dim sql As String
Dim i As Integer
If Me.TxtModFlg.text = "E" Then
Cn.Execute "Delete from TblVacationSalary where VacationID=" & val(XPTxtID.text) & ""
End If
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from TblVacationSalary where 1=-1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Me.Grid1
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
rs2.AddNew
rs2("VacationID").value = val(XPTxtID.text)
rs2("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
rs2("Emp_ID").value = val(.TextMatrix(i, .ColIndex("Emp_ID")))
rs2("RecordDate").value = IIf(.TextMatrix(i, .ColIndex("RecordDate")) = "", Null, .TextMatrix(i, .ColIndex("RecordDate")))
rs2.update
End If
Next i
End With
End Sub
Sub FillGridSalarDat()
Dim sql As String
Dim rs2 As ADODB.Recordset
Dim i As Integer
Set rs2 = New ADODB.Recordset

  With Me.Grid1
        .Clear flexClearScrollable
        .rows = 2
  End With
sql = " SELECT     dbo.TblVacationSalary.VacationID, dbo.TblVacationSalary.NetValue, dbo.TblVacationSalary.RecordDate, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode,"
sql = sql & "                       dbo.TblEmployee.Emp_Namee ,dbo.TblVacationSalary.Emp_ID"
sql = sql & " FROM         dbo.TblVacationSalary INNER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.TblVacationSalary.Emp_ID = dbo.TblEmployee.Emp_ID"
sql = sql & " Where (dbo.TblVacationSalary.VacationID = " & val(XPTxtID.text) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
  With Me.Grid1
        rs2.MoveFirst
        .rows = rs2.RecordCount + 1
     For i = 1 To rs2.RecordCount
     .TextMatrix(i, .ColIndex("payed")) = 1
     .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode"))
     .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs2("Emp_ID").value), "", rs2("Emp_ID"))
     .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(rs2("NetValue").value), 0, rs2("NetValue"))
     .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs2("RecordDate").value), "", rs2("RecordDate"))
     If SystemOptions.UserInterface = ArabicInterface Then
     .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Name").value), "", rs2("Emp_Name"))
     Else
     .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Namee").value), "", rs2("Emp_Namee"))
     End If
     rs2.MoveNext
     Next i
  End With
  End If
End Sub




Private Sub Command5_Click()
Dim StrSQL As String
              If chkGE.value = xtpChecked Then
          If ChekClodePeriod(XPDtbTrans.value) = True Then
                             If SystemOptions.UserInterface = ArabicInterface Then
                                      MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                             Else
                                  MsgBox "Please Change Date Becouse This is Period is Closed"
                            End If
                                    
                                    Exit Sub
                        End If
        Else
              
          If ChekClodePeriod(DateSta.value) = True Then
                                 If SystemOptions.UserInterface = ArabicInterface Then
                                  MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «· ”ÊÌ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                                 Else
                                 MsgBox "Please Change Date Becouse This is Period is Closed"
                                End If
              Exit Sub
         End If
     End If
                
If TxtNoteSerial.text = "" Then
CheckAccounts
createVoucher
    If Opt(0).value = True Then
           StrSQL = "Update TblEmployee Set  jopstatusid=" & Me.GetHobStatus() & ",workstate=0 Where Emp_ID=" & val(DcboEmpName.BoundText) & ""
              Cn.Execute StrSQL, , adExecuteNoRecords
      ElseIf Opt(1).value = True Then
                 StrSQL = "Update TblEmployee Set  jopstatusid=1,workstate=1 Where Emp_ID=" & val(DcboEmpName.BoundText) & ""
              Cn.Execute StrSQL, , adExecuteNoRecords
         End If
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ «·ﬁÌœ"
        Else
            MsgBox "Done"
        End If
End If
End Sub
Function CheckAccounts() As Boolean
CheckAccounts = True
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
    Dim ColumnName As String
    Dim showinMosirVac(40) As Boolean
    Dim i As Integer
    sql = "select * from mofrad order by id  "
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
 
        For i = 1 To rs.RecordCount
            FixedOrChanged(i) = IIf(IsNull(rs("FixedOrChanged").value), 0, rs("FixedOrChanged").value)
            AddOrDiscount(i) = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
            ViewComp(i) = IIf(IsNull(rs("ViewComp").value), False, rs("ViewComp").value)
            Account_code(i) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
            Account_code1(i) = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
            showMofradAll(i) = IIf(IsNull(rs("showMofradAll").value), False, rs("showMofradAll").value)
            culc30orRminder(i) = IIf(IsNull(rs("culc30orRminder").value), 0, rs("culc30orRminder").value)
            showinMosirVac(i) = IIf(IsNull(rs("showinMosirVac").value), False, rs("showinMosirVac").value)
      '      If Account_Code(i) = "" Then
      ''      MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & ViewComp(i), vbCritical
       '     getTitlesName = False
       '     Exit Function
       '     End If
            
            
            ZmamAccount(i) = IIf(IsNull(rs("ZmamAccount").value), 0, rs("ZmamAccount").value)
            AdvPaymentdAccount(i) = IIf(IsNull(rs("AdvPaymentdAccount").value), 0, rs("AdvPaymentdAccount").value)
            
            
    
              'AdvPaymentdAccount
            If SystemOptions.UserInterface = ArabicInterface Then
                componentname(i) = IIf(IsNull(rs("name").value), "", rs("name").value)
            Else
                componentname(i) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
            End If
             
              
            If ViewComp(i) = True And Account_code(i) = "" And (ZmamAccount(i) <> "True" And AdvPaymentdAccount(i) <> "True") Then
            MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & componentname(i), vbCritical
            CheckAccounts = False
          
           ' Unload Me
              Exit Function
            End If
          
             
              
         If SystemOptions.ProjectEmployeeGV = True And SystemOptions.ProjectDiscountPolicy = 1 Then 'xxx
                  If ViewComp(i) = True And AddOrDiscount(i) = -1 And Account_code1(i) = "" And (ZmamAccount(i) <> "True" And AdvPaymentdAccount(i) <> "True") Then
                MsgBox " ·„ Ì „ —»ÿ Õ”«» «·«Ì—«œ«  «· Ì  ⁄·Ì «·Œ’„ «·Œ«’ » " & componentname(i), vbCritical
        '        CheckAccounts = False
                
                '  Unload Me
                    Exit Function
                  End If
              
             End If
             
             
            rs.MoveNext
             
        Next i
  
    End If
 
    rs.Close
End Function
Private Sub Command8_Click()
Dim StrTempAccountCode As String
            Dim FirstPeriod As Date
            getFirstPeriodDateInthisYear FirstPeriod
                   'StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcCustomer.BoundText))
 
       StrTempAccountCode = get_EMPLOYEE_Account(val(DcboEmpName.BoundText), "Account_Code1")    '«·«ÃÊ— «·„” Õﬁ…
            '    StrAccountCode = Employee_account
         
         
            ShowReport StrTempAccountCode, DcboEmpName.text, FirstPeriod, Date


End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.text, , 200
End Sub

Private Sub DateSta_Change()
    If val(DcboEmpName.BoundText) = 0 Then Exit Sub
DcboEmpName_Click (0)
dateval
ChekVacation
ShowComponent
End Sub



Private Function GetVacDay() As Integer
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    s = " SELECT"
    s = s & "         SUM(dbo.TblChangedComponentRegisterDetails.NoofDays) As SumNoofDays"
    s = s & " From dbo.TblChangedComponentRegister"
    s = s & " LEFT OUTER JOIN dbo.TblChangedComponentRegisterDetails"
    s = s & "     ON dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid"
    s = s & " WHERE (dbo.TblChangedComponentRegister.Actualmonth <= " & Month(stratDate.value) & ")"
    s = s & " AND (dbo.TblChangedComponentRegister.Actualyear = " & year(stratDate.value) & ")"
    s = s & " AND ISNULL(value, 0) = 0"
    s = s & " GROUP BY dbo.TblChangedComponentRegisterDetails.Emp_id"
    s = s & " HAVING (dbo.TblChangedComponentRegisterDetails.Emp_id = " & val(DcboEmpName.BoundText) & " )"
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If Not rsDummy.EOF Then
        GetVacDay = val(rsDummy!SumNoofDays & "")
    End If
    
End Function
Private Sub DcboEmpName_Change()
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub
       TxtPreSalary.text = 0
       
   If SystemOptions.VacstionShowOldSalaries = False Then
 
                      If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
                                ClaCulte
                      Else
                               FillGridSalarDat
                       End If
       End If

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
   If Me.TxtModFlg = "R" Then Exit Sub
    Dim StrSQL As String
        Dim IssueDate As Date
        Dim DepID As Double
        Dim specid As Double
        Dim JobTypeID As Double
        Dim gradeID As Double
        Dim Account_code2 As String
           Dim Account_code  As String
        Dim Balance As String
        Dim projectId As Integer
 Dim endiqama As String
        Dim national As String
        Dim endContractPerMonth As Double
       Dim BignDateWork As Date
       Dim JobTypeName As String
       Dim JobTypeIDIQ As Integer
       Dim Contract_period As Integer
     Dim Contract_periodno As Integer
   Dim dcjopstatus As Integer
Dim LastDate As Date
Dim CountDay As Integer
Dim CountDaysal As Double
Dim BalanceCountDay As Integer
Dim BalanceVOcation As Integer
Dim cont As Integer
Dim netDay As Double
Dim TotalAbs As Double
Dim TotalWithout As Double
Dim Tiket As Double
Dim IDees As String
Dim NoVation As Double

'TxtWithOutSala1
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth, national, , , projectId, , , , , endiqama, , BignDateWork, LastDate, JobTypeName, Contract_period, Contract_periodno, , dcjopstatus, JobTypeIDIQ, , , , BalanceVOcation
        
            DcbDept.BoundText = DepID
        DcboJobsType.BoundText = JobTypeID
 lbl(33).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
        BignDate.value = BignDateWork
GetEmployeeSalaryAccordingToComponentEndserviceCotract val(DcboEmpName.BoundText)
'GetEmployeeSalaryAccordingToComponentEndservice val(DcboEmpName.BoundText)
RtriverAsse val(DcboEmpName.BoundText)
LastVocatinDate = GETlASTiSSUEDATENew((val(DcboEmpName.BoundText)), , 1)


If Opt(0).value = True Then
TxtDaySalary.text = day(DateSta.value)
Else
TxtDaySalary.text = 0
End If
ChekVacation
GetHoldayDays val(DcboEmpName.BoundText), CountDay, CountDaysal, Tiket, netDay

TxtValueTickt.text = Round(Tiket, 2)

IDes.text = GetIDesUnpadiVacation(val(DcboEmpName.BoundText))
NoVation = GetNoDayUnpadiVacation2(val(DcboEmpName.BoundText), 0)

'NoVation = 17

'TxtWithOutSala2
TxtDiscouDay.text = GetNoDayUnpadiVacation2(val(DcboEmpName.BoundText), 1)
'GetNoDayUnpadiVacation val(DcboEmpName.BoundText), IDees, NoVation
 TxtNoVaction.text = NoVation
 TxtWithOutSala1 = NoVation

 TxtVSa.text = (val(TotalWithout)) + val(TxtNoVaction.text)
 
 If val(TxtVSa.text) = 0 Then
    TxtVSa.text = GetVacDay
End If
 TxtWithOutSala1 = TxtVSa.text
 'TxtWithOutSala12 = TxtVSa.text
 TxtToalAbsent.text = Round(netDay * val(val(Me.TxtWithOutSala1.text) + val(TxtNewAbsent.text)), 2)
 
If CheckSettingsVacType() = False Then

'TxtContDay.text = CountDay
GetholidayInformationnew val(DcboEmpName.BoundText), DateSta.value
 TxtToalAbsent.text = Round(netDay * val(val(Me.TxtWithOutSala1.text) + val(TxtNewAbsent.text)), 2)
 'TxtToalAbsent.text = (TxtToalAbsent.text

'TxtDaySalVocation.text = CountDaysal
GeBalancetHoldayDays val(DcboEmpName.BoundText), BalanceCountDay

'GetholidayInformationnew
'TxtLastDayVoc.text = BalanceCountDay + BalanceVOcation
 
 GetAbsenece val(DcboEmpName.BoundText), 0, TotalWithout
 If (val(TotalWithout) / 360) >= 1 Then
 TxtYaerOut.text = (val(TotalWithout) / 360)
 TotalWithout = TotalWithout - (360 * val(TxtYaerOut.text))
 End If
  If (val(TotalWithout) / 30) >= 1 Then
 TxtMontOut.text = (val(TotalWithout) / 360)
 TotalWithout = TotalWithout - (30 * val(TxtMontOut.text))
 End If
 
 TxtVSa.text = (val(TotalWithout)) + val(TxtNoVaction.text)
 ''''''''''''''''''''
 GetAbsenece val(DcboEmpName.BoundText), 1, TotalAbs
  If (val(TotalAbs) / 360) >= 1 Then
 TxtYearAbs.text = (val(TotalAbs) / 360)
 TotalAbs = TotalAbs - (360 * val(TxtYearAbs.text))
 End If
   If (val(TotalAbs) / 30) >= 1 Then
 TxtMoAbs.text = (val(TotalAbs) / 30)
 TotalAbs = TotalAbs - (30 * val(TxtMoAbs.text))
 End If
 TxtDayAbs.text = TotalAbs
 End If
 TxtDuVocation.text = val(TxtContDay.text) - val(TxtToalAbsent.text) - val(TxtWithOutSala1.text)
 TxtTotalDay.text = val(TxtDuVocation.text) + val(TxtLastDayVoc.text)
 TxtDaySalVocation.text = val(TxtTotalDay.text)
 GetProInsurance
 TxtGetInsurance.text = GetInsurnceValue()

End Sub

Sub dateval()
If Me.TxtModFlg.text <> "R" Then
 
   Dim astrSplitItems() As String
    Dim Result As String
     Dim diff_year As Integer
    Result = ExactAge(BignDate.value, DateSta.value)
 If Result <> "" Then
    astrSplitItems = Split(Result, "-")
    TxtYear.text = astrSplitItems(0)
    TxtMonth.text = astrSplitItems(1)
    TxtDay.text = astrSplitItems(2)
  End If
      Result = ExactAge(LastVocatinDate.value, DateSta.value)
 If Result <> "" Then
    astrSplitItems = Split(Result, "-")
    TxtYear2.text = astrSplitItems(0)
    TxtMonth2.text = astrSplitItems(1)
    TxtDay2.text = astrSplitItems(2)
  End If
  TxtAddDay.text = GetLastBalanceMonthVaction(val(DcboEmpName.BoundText), val(XPTxtID.text))
  DTPicker4.value = DateAdd("d", val(TxtDiscouDay.text), LastVocatinDate)
  DTPicker3.value = DateAdd("d", val(TxtAddDay.text) * 30, DateSta)
        Result = ExactAge(DTPicker4.value, DTPicker3.value)
 If Result <> "" Then
    astrSplitItems = Split(Result, "-")
    TxtYear3.text = astrSplitItems(0)
    TxtMonth3.text = astrSplitItems(1)
    TxtDay3.text = astrSplitItems(2)
  End If
End If
End Sub


Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
'FillGridWithData3
dateval
DcboEmpName_Change
DoEvents
Dim str As String

ShowComponent

End Sub





Sub GeBalancetHoldayDays(Optional EmpID As Integer = 0, Optional ByRef NoDays As Integer)
  Dim sql As String
  Dim rs As New ADODB.Recordset
  Dim i As Integer
   Dim NODiffDate As Integer
   NODiffDate = 0
  sql = "SELECT    * from TblEmpHolidaysDetails WHERE     (Emp_id = " & EmpID & ")"
  rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
  If rs.RecordCount > 0 Then
  For i = 1 To rs.RecordCount
If Not IsNull(rs("DateExpectedM").value) Then
                If Not IsNull(rs("todate").value) Then
                 NODiffDate = NODiffDate + (val(DateDiff("d", rs("DateExpectedM").value, rs("todate").value)) * -1)
                 End If
                 End If
               rs.MoveNext
Next i
Else
End If
NoDays = NODiffDate
       
End Sub
Sub GetHoldayDays(Optional EmpID As Integer = 0, Optional ByRef NoDays As Integer, Optional ByRef NoDaysSala As Double, Optional ByRef Tiket As Double, Optional ByRef netDay As Double)
  Dim sql As String
  Dim rs As New ADODB.Recordset
  Dim HoldaType As Integer
  Dim HoldaNo As Double
  Dim PriodType As Integer
  Dim PriodNo As Double
  Dim PriodType2 As Integer
  Dim PriodNo2 As Double
  Dim NODiffDate As Integer
  Dim tktval As Double
  Dim tkno As Integer
  sql = "SELECT    * from dbo.Contract WHERE     (Emp_id = " & EmpID & ")"
  rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
  If rs.RecordCount > 0 Then
 HoldaNo = IIf(IsNull(rs("Holiday_period_no").value), 0, rs("Holiday_period_no").value)
 HoldaType = IIf(IsNull(rs("Holiday_period").value), -1, rs("Holiday_period").value)
 PriodNo = IIf(IsNull(rs("Due_period_no").value), 0, rs("Due_period_no").value)
 PriodType = IIf(IsNull(rs("due_period").value), -1, rs("due_period").value)
  PriodNo2 = IIf(IsNull(rs("salary_period_no").value), 0, rs("salary_period_no").value)
 PriodType2 = IIf(IsNull(rs("salary_period").value), -1, rs("salary_period").value)
tkno = IIf(IsNull(rs("no_of_Child_ticket").value), 0, rs("no_of_Child_ticket").value)
    tktval = IIf(IsNull(rs("TicketValue").value), 0, rs("TicketValue").value)
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
PriodNo = PriodNo * 360
End If

NODiffDate = DateDiff("d", LastVocatinDate.value, DateSta.value)
If NODiffDate >= PriodNo Then
'If val(NODiffDate / 360) <= 1 Then
NoDaysSala = (PriodNo2 / PriodNo) * PriodNo
NoDays = (HoldaNo / PriodNo) * PriodNo
netDay = (HoldaNo / PriodNo)
'ElseIf val((NODiffDate) / 360) < 2 Then
'NoDays = (HoldaNo / PriodNo) * 360
'NoDaysSala = (PriodNo2 / PriodNo) * 360
'Else
'NoDays = (HoldaNo / PriodNo) * 720
'NoDaysSala = (PriodNo2 / PriodNo) * 720
End If
Else
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
 Sub GetEmployeeSalaryAccordingToComponentEndserviceCotract(Emp_id As Integer)
                                                    
    Dim sql As String
    Dim mofrad_name As String
    Dim valuee As Double
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim Mofradd As String
    Dim i As Integer
    With Me.Fg
    .rows = 1
    End With
sql = " SELECT     dbo.mofrad.name, dbo.mofrad.nameE, dbo.TblContractDetails.Mofradtype, dbo.TblContractDetails.Emp_id"
sql = sql & " FROM         dbo.TblContractDetails INNER JOIN"
sql = sql & "                      dbo.mofrad ON dbo.TblContractDetails.Mofradtype = dbo.mofrad.id"
sql = sql & " WHERE     (dbo.TblContractDetails.Emp_id = " & Emp_id & ") AND (dbo.TblContractDetails.DefDataType = 0)"

      rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

      For i = 1 To rs.RecordCount
   GetEmployeeSalaryAccordingToComponentEndservice Emp_id, val(IIf(IsNull(rs("Mofradtype").value), 0, rs("Mofradtype").value))
  
 rs.MoveNext
      Next i

     End If
      rs.Close
  
End Sub
 Sub GetEmployeeSalaryAccordingToComponentEndservice(Emp_id As Integer, Optional MofrID As Integer = 0)
                                                    
    Dim sql As String
    Dim mofrad_name As String
    Dim valuee As Double
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim Mofradd As String
    Dim i, j As Integer
    

sql = "    SELECT     TOP 100 PERCENT dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.mofrdat.mofrad_type, dbo.mofrad.AddOrDiscount,"
sql = sql & "                      SUM(dbo.EmpSalaryComponent.[Value]) AS SmValue, dbo.EmpSalaryComponent.AccountCode, dbo.EmpSalaryComponent.mofrad_type AS mofrad_typeDet"
sql = sql & "   FROM         dbo.mofrad INNER JOIN"
sql = sql & "                        dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type RIGHT OUTER JOIN"
sql = sql & "                        dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode"
sql = sql & "   GROUP BY dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.mofrdat.mofrad_type, dbo.mofrad.AddOrDiscount, dbo.EmpSalaryComponent.emp_ID,"
sql = sql & "                        dbo.EmpSalaryComponent.AccountCode , dbo.EmpSalaryComponent.mofrad_type"
sql = sql & "  HAVING      (dbo.EmpSalaryComponent.emp_ID = " & Emp_id & ") AND (dbo.mofrdat.mofrad_type = " & MofrID & ")"
sql = sql & "   ORDER BY dbo.mofrdat.mofrad_type"
      rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
  With Me.Fg
 ' .Rows = rs.RecordCount + 1
j = .rows
.rows = .rows + rs.RecordCount

        For i = j To .rows - 1
     
       .TextMatrix(i, .ColIndex("Serial")) = i
      .TextMatrix(i, .ColIndex("mofrdID")) = IIf(IsNull(rs("mofrad_type").value), 0, rs("mofrad_type").value)
      If SystemOptions.UserInterface = ArabicInterface Then
       .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
       Else
       .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(rs("mofrad_namee").value), "", rs("mofrad_namee").value)
       End If
 .TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(rs("SmValue").value), 0, rs("SmValue").value)
  
 rs.MoveNext
      Next i
 End With
     End If
      rs.Close
    ReLineGrid
End Sub

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter  As Integer
   lbl(10).Caption = 0
    IntCounter = 0

    With Fg

        For i = .FixedRows To .rows - 1

       
 If val(.TextMatrix(i, .ColIndex("Valu"))) > 0 Then
                
                lbl(10).Caption = val(lbl(10).Caption) + val(.TextMatrix(i, .ColIndex("Valu")))
        
            End If
        Next i
 
    End With
    


End Sub

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
  FrmEmployeeSearch.lbltype = 26
      Set FrmEmployeeSearch.RetrunFrm = Me

      FrmEmployeeSearch.show
  
    End If
End Sub

Private Sub Dcbranch_Change()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.text = ""
End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.text = ""
End If

End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DISACC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 6661
    End If


End Sub

Private Sub ENDDATE_Change()
If Me.TxtModFlg.text <> "R" Then
TxtNoMonth.text = DateDiff("M", stratDate.value, EndDate.value)
GetProInsurance
 TxtGetInsurance.text = GetInsurnceValue()
End If
End Sub



Private Sub Grid1_Click()
RelinSalaryPayed
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub
Sub RelinSalaryPayed()
Dim SumSalary As Double
SumSalary = 0
Dim i As Integer
With Grid1
For i = 1 To .rows - 1
If Grid1.cell(flexcpChecked, i, Grid1.ColIndex("payed")) = flexChecked Then
If val(.TextMatrix(i, .ColIndex("NetValue"))) <> 0 Then
SumSalary = SumSalary + val(.TextMatrix(i, .ColIndex("NetValue")))
End If
End If
Next i
End With
TxtPreSalary.text = SumSalary
Ch(8).value = vbChecked
End Sub
Function CHeckPayedSalaryCurrMonth() As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " select RecordDate from emp_salary"
sql = sql + " where dbo.emp_salary.emp_id=" & val(Me.DcboEmpName.BoundText)
sql = sql + " AND (payed =1 )  AND     (sgn = '" & year(DateSta.value) & Month(DateSta.value) & "') "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CHeckPayedSalaryCurrMonth = True
Else
CHeckPayedSalaryCurrMonth = False
End If
End Function
Sub ClaCulte()
'Exit Sub
Dim FurstDate As Date
Dim RecDate As Date
Dim TempDate As Date
Dim Diff As Double
Dim i As Integer
'Diff = 0
'FurstDate = GetMaxDatSalaryPayed()
'TempDate = FurstDate
'Diff = DateDiff("M", FurstDate, DateSta.value)
'  With Me.Grid1
'        .Clear flexClearScrollable
'        .Rows = 1
'  End With
'If Diff > 0 Then
'  With Me.Grid1
'        .Clear flexClearScrollable
'        .Rows = Diff + 1
'  End With
'  Dim TempDif  As Double
'  TempDif = Diff
'  If Diff > 1 Then
'  Diff = Diff - 1
'  End If
'For i = 1 To Diff
'If TempDif <> 1 Then
'RecDate = DateAdd("m", 1, TempDate)
'Else
'RecDate = FurstDate
'End If
'TempDate = RecDate
'FillGridWithDataSalary i, RecDate
'Next i
'End If
Dim str As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
If val(DcboEmpName.BoundText) = 0 Or DcboEmpName.text = "" Then Exit Sub
sql = " select RecordDate from emp_salary"
sql = sql + " where dbo.emp_salary.emp_id=" & val(Me.DcboEmpName.BoundText)
If Me.TxtModFlg.text = "N" Then
sql = sql + " AND (payed =0  or payed is null )  AND     (sgn <> '" & year(DateSta.value) & Month(DateSta.value) & "') and not (RecordDate is null) "
ElseIf Me.TxtModFlg.text = "E" Then
sql = sql + " AND (payed =0  or payed is null  or VocEntitID =" & val(XPTxtID.text) & " )  AND     (sgn <> '" & year(DateSta.value) & Month(DateSta.value) & "') and not (RecordDate is null) "
End If
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  With Me.Grid1
        .Clear flexClearScrollable
        .rows = 1
  End With
If rs2.RecordCount > 0 Then
  With Me.Grid1
        .Clear flexClearScrollable
        .rows = rs2.RecordCount + 1
  End With
rs2.MoveFirst
str = "01/01/1991"
DTPicker5.value = CDate(str)
For i = 1 To rs2.RecordCount
RecDate = IIf(IsNull(rs2("RecordDate").value), DTPicker5, rs2("RecordDate").value)
If DTPicker5.value <> RecDate Then
FillGridWithDataSalary i, RecDate
End If
rs2.MoveNext
Next i
End If
RelinSalaryPayed
End Sub
Function GetMaxDatSalaryPayed2() As Date
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
If val(DcboEmpName.BoundText) = 0 Or DcboEmpName.text = "" Then Exit Function
sql = " select MIN(RecordDate) as RecordDate from emp_salary"
sql = sql + " where dbo.emp_salary.emp_id=" & val(Me.DcboEmpName.BoundText)
sql = sql + "   AND     (sgn <> '" & year(DateSta.value) & Month(DateSta.value) & "') and not (RecordDate is null) "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxDatSalaryPayed2 = IIf(IsNull(rs2("RecordDate").value), LastVocatinDate.value, rs2("RecordDate").value)
Else
GetMaxDatSalaryPayed2 = LastVocatinDate.value
End If
End Function

Function GetMaxDatSalaryPayed() As Date
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
If val(DcboEmpName.BoundText) = 0 Or DcboEmpName.text = "" Then Exit Function
sql = " select max(RecordDate) as RecordDate from emp_salary"
sql = sql + " where dbo.emp_salary.emp_id=" & val(Me.DcboEmpName.BoundText)
sql = sql + " AND (payed =1)  AND     (sgn <> '" & year(DateSta.value) & Month(DateSta.value) & "') and not (RecordDate is null) "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMaxDatSalaryPayed = IIf(IsNull(rs2("RecordDate").value), GetMaxDatSalaryPayed2, DateAdd("m", 1, rs2("RecordDate").value))
Else
GetMaxDatSalaryPayed = GetMaxDatSalaryPayed2
End If
End Function
Public Sub FillGridWithData3()
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
    Dim j As Integer
    Dim ColumnName As String
    Dim netvalue As Double
    Dim OldValue As Double
    Dim RemainValue As Double
    Dim PaymentValue As Double
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    With Me.Grid1
            .rows = 2
            .Clear flexClearScrollable
    End With
If val(DcboEmpName.BoundText) = 0 Or DcboEmpName.text = "" Then Exit Sub
My_SQL = " SELECT   TblEmployee.BranchId AS EMPBRANCHID,  *, dbo.EmpGroupDep.GroupName AS GroupName1, dbo.EmpGroupDep.Ename AS Ename1, dbo.TblEmpDepartments.DepartmentName AS DepartmentName1,"
My_SQL = My_SQL + "                       dbo.TblEmpDepartments.DepartmentNamee AS DepartmentNamee1, dbo.projects.Project_name AS Project_name1, dbo.projects.Project_nameE AS Project_nameE1,"
My_SQL = My_SQL + "                       dbo.emp_contract_type.name AS name1, dbo.emp_contract_type.NameE AS NameE1 , dbo.emp_salary.id AS IDEmp ,dbo.TblEmployee.SalaryCode"
My_SQL = My_SQL + "  FROM         dbo.emp_salary INNER JOIN"
My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.emp_salary.emp_id = dbo.TblEmployee.Emp_ID INNER JOIN"
My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.emp_salary.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
My_SQL = My_SQL + "                       dbo.emp_contract_type ON dbo.TblEmployee.ContractID = dbo.emp_contract_type.id LEFT OUTER JOIN"
My_SQL = My_SQL + "                       dbo.projects ON dbo.emp_salary.project_id = dbo.projects.id LEFT OUTER JOIN"
My_SQL = My_SQL + "                       dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
My_SQL = My_SQL + "                       dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID"
My_SQL = My_SQL + "                       "
My_SQL = My_SQL + "   WHERE     ( 1=1) "
My_SQL = My_SQL + " and dbo.emp_salary.emp_id=" & val(Me.DcboEmpName.BoundText)
If Me.TxtModFlg.text = "N" Then
My_SQL = My_SQL + " AND (payed =0)  AND     (sgn <> '" & year(DateSta.value) & Month(DateSta.value) & "')  "
End If
If Me.TxtModFlg.text = "R" Then
My_SQL = My_SQL + " AND VocEntitID=" & val(Me.XPTxtID.text) & "  and not (VocEntitID is null)"
End If
If Me.TxtModFlg.text = "E" Then
My_SQL = My_SQL + " AND (((payed =0)  AND     (sgn <> '" & year(DateSta.value) & Month(DateSta.value) & "')) or ( VocEntitID=" & val(Me.XPTxtID.text) & " and not (VocEntitID is null) ))  "
End If
My_SQL = My_SQL + " ORDER BY dbo.emp_salary.RecordDate"

Dim k As Integer
  Dim Emp_id As Double
   
        rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        With Me.Grid1
            .rows = 2
            .Clear flexClearScrollable

            If rs.RecordCount > 0 Then
               ' .Rows = rs.RecordCount + 1
               .rows = 1
                rs.MoveFirst
i = 0
                For k = 1 To rs.RecordCount
                OldValue = 0
                netvalue = 0
                      netvalue = IIf(IsNull(rs.Fields("EmpTotalNet").value), 0, Round(rs.Fields("EmpTotalNet").value, 2))
                      Emp_id = IIf(IsNull(rs.Fields("Emp_id").value), 0, rs.Fields("Emp_id").value)
           
                    
                  If netvalue <> OldValue And netvalue <> 0 Then
                  .rows = .rows + 1
                    i = i + 1
                    .TextMatrix(i, .ColIndex("Ser")) = i
                    .TextMatrix(i, .ColIndex("NetValue")) = netvalue
             .TextMatrix(i, .ColIndex("payed")) = IIf(IsNull(rs.Fields("payed").value), 0, rs.Fields("payed").value)
              
                        If .TextMatrix(i, .ColIndex("payed")) = True Then
                .cell(flexcpBackColor, i, 1, i, 62) = &HFF00&
            Else
                .cell(flexcpBackColor, i, 1, i, 62) = vbWhite
            End If
            .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("IDEmp").value), "", rs.Fields("IDEmp").value)
            .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("EMPBRANCHID").value), "", rs.Fields("EMPBRANCHID").value)
           .TextMatrix(i, .ColIndex("Emp_id")) = Emp_id
                    
                    .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                    .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs.Fields("cost_center_id").value), "", rs.Fields("cost_center_id").value)
            
                    .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                    .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
                      Dim str As String
   ' str = "01/" & CmbMonth.ListIndex + 1 & "/" & CboYear.Text

   ' DTPicker2.value = MonthLastDay(CDate(str))
   ' DTPicker2 = MonthLastDay(CDate(str))
  '  .TextMatrix(i, .ColIndex("RecordDate")) = DTPicker2.value
                    .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs.Fields("RecordDate").value), "", rs.Fields("RecordDate").value)
                    .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
                  If SystemOptions.UserInterface = ArabicInterface Then
                  .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value)
                  .TextMatrix(i, .ColIndex("GroupName1")) = IIf(IsNull(rs.Fields("GroupName1").value), "", rs.Fields("GroupName1").value)
                  .TextMatrix(i, .ColIndex("DepartmentName1")) = IIf(IsNull(rs.Fields("DepartmentName1").value), "", rs.Fields("DepartmentName1").value)
                  .TextMatrix(i, .ColIndex("Project_name1")) = IIf(IsNull(rs.Fields("Project_name1").value), "", rs.Fields("Project_name1").value)
                  .TextMatrix(i, .ColIndex("name1")) = IIf(IsNull(rs.Fields("name1").value), "", rs.Fields("name1").value)
                  .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value)
                  
                  Else
                  .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs.Fields("branch_name").value), "", rs.Fields("branch_name").value)
                  .TextMatrix(i, .ColIndex("GroupName1")) = IIf(IsNull(rs.Fields("Ename1").value), "", rs.Fields("Ename1").value)
                  .TextMatrix(i, .ColIndex("DepartmentName1")) = IIf(IsNull(rs.Fields("DepartmentNamee1").value), "", rs.Fields("DepartmentNamee1").value)
                  .TextMatrix(i, .ColIndex("Project_name1")) = IIf(IsNull(rs.Fields("Project_nameE1").value), "", rs.Fields("Project_nameE1").value)
                  .TextMatrix(i, .ColIndex("name1")) = IIf(IsNull(rs.Fields("NameE1").value), "", rs.Fields("NameE1").value)
                  End If
                    .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Round(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                    .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("Mokafea").value), "", Round(rs.Fields("Mokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                    .TextMatrix(i, .ColIndex("TotalAdvance")) = IIf(IsNull(rs.Fields("TotalAdvance").value), "", Round(rs.Fields("TotalAdvance").value))
            
                    .TextMatrix(i, .ColIndex("SalesCom")) = IIf(IsNull(rs.Fields("SalesCom").value), "", Round(rs.Fields("SalesCom").value))
                    
                    .TextMatrix(i, .ColIndex("total1")) = IIf(IsNull(rs.Fields("total1").value), "", Round(rs.Fields("total1").value, 2))
            
                    .TextMatrix(i, .ColIndex("total2")) = IIf(IsNull(rs.Fields("total2").value), "", Round(rs.Fields("total2").value, 2))
                    .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(rs.Fields("EmpTotalNet").value), "", Round(rs.Fields("EmpTotalNet").value, 2))

                    For j = 1 To 40
                       ColumnName = "Comp" & j
                       .TextMatrix(i, .ColIndex(ColumnName)) = IIf(IsNull(rs.Fields("EmpTotalNet").value), "", Format(rs.Fields(ColumnName).value))
                     Next j
                    
            End If
            rs.MoveNext
                Next k

                rs.Close
            End If
    
          '  GetAdvanceValues IntMonth, IntYear
          '  GetWorkHours
          '  CalculateNets
            .rows = .rows + 1

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
            Else
                .TextMatrix(.rows - 1, .ColIndex("Ser")) = "Total"
            End If

            .IsSubtotal(.rows - 1) = True
            Dim SngTotal As Single
      
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows - 1, .ColIndex("EmpTotalNet"))
            .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
            net_value1 = SngTotal
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .rows - 1, .ColIndex("CorrectEmpTotalNet"))
            .TextMatrix(.rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal

    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
            .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
            .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
            .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
            .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
            '               SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
            .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
    
            '               SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
            .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
        End With

     

Set rs = Nothing
Check21.value = vbChecked
Check21_Click
   RelinSalaryPayed
ErrTrap:
End Sub

Private Sub Check21_Click()
    Dim i As Integer

    If Check21.value = vbChecked Then

        With Me.Grid1
 
            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.Grid1

            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If
    
RelinSalaryPayed
     End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim My_SQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic

                  If SystemOptions.SpecialVersion = True Then
     Ele(11).Visible = False
    
End If


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
    Dcombos.GetUsers Opretot
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetJobEndService dctype
  
 Dcombos.GetAccountingCodes ADDACC, True
Dcombos.GetAccountingCodes DISACC, True

  
    My_SQL = "select Emp_id,Emp_Name From TblEmployee  order by  Emp_Name"
    fill_combo Dcemp, My_SQL

  
 
    With Me.Grid
        .rows = 1
        .Clear flexClearScrollable
    End With

 

     If getTitlesName = True Then
   
   End If
   
    
     YearMonth
     
     
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEmpDepartments Me.DcbDept
    Dcombos.GetEmpJobsTypes Me.DcboJobsType

    If SystemOptions.UserInterface = ArabicInterface Then
    CbBasedOn.AddItem "»·«"
    CbBasedOn.AddItem "ÿ·» ≈Ã«“…"
    Else
    CbBasedOn.AddItem "Without"
    CbBasedOn.AddItem "Vacation Request"
    End If
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If
    
    SetDtpickerDate Me.XPDtbTrans
  '  YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblVocationEntitlements   WHERE     (Flag IS NULL) Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
        Me.TxtModFlg.text = "R"
        DB_CreateField "TblVocationEntitlements", "PaymentRecommended", adDouble, adColNullable, , , "      ", False, True
    Retrive


    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    Option1.value = True
    Exit Sub
    ' Chekk
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
   lbl(32).Caption = "Salary"
   lbl(50).Caption = "Unpaid Vac"
   lbl(51).Caption = "Net Absnce"
   Command3.Caption = "Show"
   lbl(63).Caption = "Pre.Salary"
   C1Tab1.Caption = "Data|Approve"
   Label1(1).Caption = "ADD ACC."
   Label1(2).Caption = "Dis ACC."
   
             Label10.Caption = "Currently required Approve"
       lbl(67).Caption = "Work Period"
    lbl(65).Caption = "Year"
    lbl(64).Caption = "Month"
    lbl(66).Caption = "Day"
    lbl(68).Caption = "Balance"
    lbl(69).Caption = "Month"
''//
    With VSFlexGrid2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
    End With
Command5.Caption = "Create GE"
Command2.Caption = "Delete GE"
lbl(55).Caption = "End Service"
Opt(2).RightToLeft = False
Opt(2).Caption = "End Service"
Opt(0).RightToLeft = False
Opt(1).RightToLeft = False
Opt(0).Caption = "Vaca."
Opt(1).Caption = "On Work"
Label3.Caption = "End"
Label2.Caption = "Start"
lbl(57).Caption = "Insu.Value"
Label5.Caption = "No Vacatio"
    lbl(44).Caption = "Day"
    lbl(42).Caption = "Month"
    lbl(43).Caption = "Year"
    lbl(45).Caption = "Day"
    lbl(47).Caption = "Month"
    lbl(46).Caption = "Year"
    Option1.Caption = "Entitlements"
    Option2.Caption = "Commitment"
    lbl(56).Caption = "Based On"
'//
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(9).Caption = "Print Entitlements"
    Cmd(8).Caption = "Print Commitment"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    Label4.Caption = "Req.No"
    Label1(35).Caption = "GL No."
    Label1(0).Caption = "Settlement"
    XPTab301.Caption = "Data"
    lbl(48).Caption = "Absence"
    lbl(70).Caption = "Deductes"
    lbl(71).Caption = "Net. Ds"
     lbl(72).Caption = "Day"
     lbl(74).Caption = "Month"
     lbl(73).Caption = "Year"
     
     
    Frame7.Caption = "Absence"
    Command1.Caption = "Show"
    Accredit.Caption = "Send Approved"
  lbl(60).Caption = "Without Salary"
    Me.Caption = "  Vacation Due"
    EleHeader.Caption = Me.Caption
    lbl(58).Caption = "Days Absence"
  '  Frame9.Caption = "Accounting"
    Command9.Caption = "Print GL"
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lbl(49).Caption = "Days leave entitlement"
    lbl(23).Caption = "Employee"
    lbl(3).Caption = "Employee"
    lblbr.Caption = "Branch"
   lbl(24).Caption = "Job"
   lbl(61).Caption = "Salary"
   Command8.Caption = "Acc.Statement"
    lbl(0).Caption = "Department"
  lbl(59).Caption = "Service"
   lbl(35).Caption = "Day"
    lbl(36).Caption = "Month"
    lbl(28).Caption = "Year"
    lbl(9).Caption = "Start Work"
    lbl(20).Caption = "Last vacation"
    lbl(31).Caption = "Remarks"
    lbl(22).Caption = "Total Day"
    lbl(2).Caption = "Day Befor Discount"
    lbl(21).Caption = "Balance of a previous vacation"
    lbl(5).Caption = "Vocation Components"
    lbl(14).Caption = "Financial Dues"
    lbl(11).Caption = "Total "
    lbl(19).Caption = "Financial Deductions"
    lbl(12).Caption = "No Days"
    lbl(13).Caption = "Amount Due"
    lbl(15).Caption = "Current Salary"
    lbl(16).Caption = "Overtime"
    lbl(17).Caption = "Vacation Salary"
    lbl(18).Caption = "Other "
    lbl(34).Caption = "Total Amount Due "
    lbl(26).Caption = "Advance Due "
    lbl(29).Caption = "Other "
    lbl(37).Caption = "Total Amount Deducted "
    lbl(38).Caption = "Net Benefits "
    lbl(39).Caption = "Tickets position "
    ChBooked.RightToLeft = False
    ChBooked.Caption = "Booked"
    ChDelivery.RightToLeft = False
    ChDelivery.Caption = "Delivery of the ticket value"
   lbl(40).Caption = "Ticket Value"
   lbl(41).Caption = " Total Dues"
   lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "Rec. count"
'Label1.Caption = "Settlement Date"
Check21.RightToLeft = False
Check21.Caption = "Select All"
   With Me.Grid1
       .TextMatrix(0, .ColIndex("Ser")) = "Serial"
       .TextMatrix(0, .ColIndex("payed")) = "Select"
       .TextMatrix(0, .ColIndex("Emp_Code")) = "Code"
       .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
       .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
       .TextMatrix(0, .ColIndex("NetValue")) = "Value"
    End With
   With Me.Fg
       .TextMatrix(0, .ColIndex("Serial")) = "Serial"
       .TextMatrix(0, .ColIndex("mofrd")) = "Component"
       .TextMatrix(0, .ColIndex("Valu")) = "Value"

    End With
       With Me.VSFlexGrid1
       .TextMatrix(0, .ColIndex("Serial")) = "Serial"
       .TextMatrix(0, .ColIndex("AsCode")) = "No"
       .TextMatrix(0, .ColIndex("mofrd")) = "Name"
         .TextMatrix(0, .ColIndex("ReciveDate")) = "ReciveDate"
       .TextMatrix(0, .ColIndex("DeliverDate")) = "DeliverDate"
       .TextMatrix(0, .ColIndex("Emp_NameTo")) = "Recipient Name"

    End With
    
    With Grid
        .TextMatrix(0, .ColIndex("dep")) = "Department"
        .TextMatrix(0, .ColIndex("branchname")) = "Branch"
        .TextMatrix(0, .ColIndex("project")) = "Project"
        .TextMatrix(0, .ColIndex("Ser")) = "No."
        .TextMatrix(0, .ColIndex("Emp_ID")) = "Employee No."
        .TextMatrix(0, .ColIndex("Emp_Code")) = "Employee Code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
        .TextMatrix(0, .ColIndex("JobTypeName")) = "Job"
        .TextMatrix(0, .ColIndex("BignDateWork")) = " First Date of Commencement of Work"
        .TextMatrix(0, .ColIndex("lastHolidaydate")) = " Last Date of Commencement of Work"
        .TextMatrix(0, .ColIndex("WorkHours")) = "Actual Work Hours"
        .TextMatrix(0, .ColIndex("OverTime")) = " Overtime"
        .TextMatrix(0, .ColIndex("DefWorkHours")) = " Official Work Hours"
        .TextMatrix(0, .ColIndex("Mokafea")) = "Bonus"
        .TextMatrix(0, .ColIndex("TotalDiscount")) = "Penalty"
        .TextMatrix(0, .ColIndex("TotalAdvance")) = "Advance"
        .TextMatrix(0, .ColIndex("total1")) = "Total Bonuses"
        .TextMatrix(0, .ColIndex("total2")) = "Total Penalties"
        .TextMatrix(0, .ColIndex("EmpTotalNet")) = "NET"
        .TextMatrix(0, .ColIndex("sgn")) = "Signature"
        .TextMatrix(0, .ColIndex("SalesCom")) = "Commissions"
    End With
    
    chkGE.Caption = "GE date is transaction date"
    

End Sub
Private Sub YearMonth()
    Dim i As Integer
    Dim IntDefIndex As Integer
    CmbMonth.Clear
    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next
    CmbMonth.ListIndex = Month(Date) - 1
    ''''''''''
    CboYear.Clear
    For i = 2000 To 2050
        CboYear.AddItem i
        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If
    Next
    CboYear.ListIndex = IntDefIndex
End Sub


' Private Sub YearMonth()

'    Dim i As Integer
'    Dim IntDefIndex As Integer

  '  CmbMonth.Clear

 '   For i = 1 To 12
    '    CmbMonth.AddItem MonthName(i)
   ' Next

   ' CmbMonth.ListIndex = Month(Date) - 1
   ' CboYear.Clear

  '  For i = 2010 To 2050
  '      CboYear.AddItem i
'
'        If i = year(Date) Then
'            IntDefIndex = CboYear.NewIndex
'        End If

'    Next

'    CboYear.ListIndex = IntDefIndex
'End Sub

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



Private Sub Label6_Click()
C1Elastic2.Visible = False
Ele(10).Visible = True
Ele(12).Visible = False
End Sub

Private Sub Label7_Click()
C1Elastic2.Visible = True
Ele(10).Visible = False
Ele(12).Visible = False
End Sub

Private Sub Label8_Click()
Ele(10).Visible = True
C1Elastic2.Visible = False
Ele(12).Visible = False
End Sub

Private Sub NetTotal_Change()
Total.text = val(NetTotal.text) + val(TxtValueTickt.text)
End Sub



Private Sub Opt_Click(index As Integer)
If Opt(2).value = True Then
dctype.Enabled = True
Else
dctype.Enabled = False
End If
DcboEmpName_Click (0)
End Sub

Private Sub stratDate_Change()
If Me.TxtModFlg.text <> "R" Then
TxtNoMonth.text = DateDiff("M", stratDate.value, EndDate.value)
GetProInsurance
 TxtGetInsurance.text = GetInsurnceValue()
End If
End Sub

Private Sub TxtAbsent_Change()
'TxtToalAbsent.text = val(Me.txtToOutSal.text) + val(TxtAbsent.text)
End Sub

Private Sub TxtAdvance_Change()
Smation
End Sub

Private Sub TxtContDay_Change()
'TxtTotalDay.text = val(TxtContDay.text) + val(TxtLastDayVoc.text) - val(TxtToalAbsent.text)
TxtDuVocation.text = val(TxtContDay.text) - val(TxtToalAbsent.text) - val(TxtWithOutSala1.text)
End Sub



Private Sub TxtDayAbs_Change()
'TxtAbsent.text = val(Me.TxtDayAbs.text) + val(Me.TxtMoAbs.text) * 30 + val(Me.TxtYearAbs.text) * 360

'TxtDuVocation.text = val(TxtContDay.text) - val(TxtToalAbsent.text)
'DcboEmpName_Click (0)
End Sub

Private Sub TxtDaySalary_Change()
If Me.TxtModFlg.text <> "R" Then
lbl(33).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
Dim Sal As Double
Sal = Round(val(lbl(33).Caption), 2)

Sal = Sal / 30


Sal = Round(Sal, 2)



Dim CountDays As Double
 
Dim MonthDayNo  As Double
CountDays = day(DateSta.value)
MonthDayNo = daysInMonth(DateSta.value)

If MonthDayNo = CountDays Then
TxtSalary.text = 30 * Sal
Else

TxtSalary.text = val(TxtDaySalary.text) * Sal
End If
TxtSalary.text = Round(val(TxtSalary.text), 2)
End If
End Sub
Function GetholidayInformationnew(EmpID As Integer, DateSta As Date)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset

 Dim i As Integer
 Dim OldValue As Double
 Dim currentvalue As Double
 
       StrSQL = "Select *   From tblVacationData Where Status1 is null  and ExpectedacationDate<= " & SQLDate(DateSta, True) & " and  EmpID='" & EmpID & "'"
    
        Set rs = New ADODB.Recordset
 
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
OldValue = 0
currentvalue = 0
   If rs.RecordCount > 0 Then
   
            For i = 1 To rs.RecordCount
                    If i < rs.RecordCount Then
                        OldValue = OldValue + IIf(IsNull(rs("value").value), 0, rs("value").value)
                    ElseIf i = rs.RecordCount Then
                  currentvalue = IIf(IsNull(rs("value").value), 0, rs("value").value)
                    End If
                    rs.MoveNext
            
            Next i
   
   Else
   
   
   End If
TxtLastDayVoc.text = OldValue
TxtContDay.text = currentvalue

        rs.Close
        Set rs = Nothing
     

End Function

Private Sub TxtDaySalVocation_Change()
Dim Sal As Double
Sal = val(lbl(10).Caption)
Sal = Sal / 30
'Sal = Round(Sal, 2)
TxtSalaryVocation.text = val(TxtDaySalVocation.text) * Sal
TxtSalaryVocation.text = Round(val(TxtSalaryVocation.text), 2)
End Sub

Private Sub TxtGetInsurance_Change()
If Me.TxtModFlg.text <> "R" Then
TxtInsuranceValue.text = val(TxtNoMonth.text) * val(TxtGetInsurance.text)
End If
End Sub

Private Sub TxtIncrease_Change()
'TxtTolaMostak.text = val(Txtsalary.text) + val(Txtincrease.text) + val(TxtSalaryVocation.text) + val(TxtSalEntitOther.text)
Smation
End Sub


Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
Command5.Enabled = False
'Accredit.Enabled = False
    Select Case Me.TxtModFlg.text

        Case "R"
'        Accredit.Enabled = True
        chkGE.Enabled = False
        Frame1.Enabled = False
        Command5.Enabled = True
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  "
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(8).Enabled = True
            Me.Cmd(9).Enabled = True
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
        chkGE.Enabled = True
        
        Frame1.Enabled = True
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(8).Enabled = False
            Me.Cmd(9).Enabled = False
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
          '  TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
        chkGE.Enabled = True
        Frame1.Enabled = True
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  (  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(8).Enabled = False
            Me.Cmd(9).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
          '  TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtMontOut_Change()
'txtToOutSal.text = val(TxtVSa.text) + val(TxtMontOut.text) * 30 + val(TxtYaerOut.text) * 360
'DcboEmpName_Click (0)
End Sub

Private Sub TxtNewAbsent_Change()
DcboEmpName_Change
End Sub

Private Sub TxtNoMonth_Change()
If Me.TxtModFlg.text <> "R" Then
TxtInsuranceValue.text = val(TxtNoMonth.text) * val(TxtGetInsurance.text)
End If
End Sub
Sub GetRequstVacation(Optional OrderID As Double = 0)
If OrderID <> 0 Then
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim Scren As String
Scren = "formvocatinl"
sql = "select * from TblVocation where ID =" & OrderID & ""
 If CheckAprroveScreen("formvocatinl") = True Then
sql = sql & " and   (dbo.ScreenSendAparoved(" & OrderID & ", '" & Scren & "') > 0)"
sql = sql & " and   (dbo.ScreenIsAparoved(" & OrderID & ", '" & Scren & "') is null)"
End If
Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
stratDate.value = IIf(IsNull(Rs7("FromDate").value), Date, Rs7("FromDate").value)
EndDate.value = IIf(IsNull(Rs7("ToDate").value), Date, Rs7("ToDate").value)
DcboEmpName.BoundText = IIf(IsNull(Rs7("EmpID").value), 0, Rs7("EmpID").value)
TxtNoVacation.text = IIf(IsNull(Rs7("NoVacation").value), 0, Rs7("NoVacation").value)
ENDDATE_Change
End If
End If
End Sub
Function CheckVacation() As Boolean
Dim sql As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "select * from TblVocationEntitlements where NoOrder  =" & val(Txtorder.text) & " and ID <> " & val(XPTxtID.text) & ""
Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
CheckVacation = True
Else
CheckVacation = False
End If
End Function

Public Sub TxtOrder_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If Me.TxtModFlg.text <> "R" Then
If val(CbBasedOn.ListIndex) = 1 Then
If CheckVacation() = False Then
GetRequstVacation val(Txtorder.text)
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ ⁄„· «” Õﬁ«ﬁ ·Â–« «·ÿ·» ”«»ﬁ«"
Else
MsgBox "It has been previously worked entitlement to this request "
End If
Exit Sub
End If
End If
End If
End If
End Sub


Private Sub TxtOrder_KeyUp(KeyCode As Integer, Shift As Integer)
If val(CbBasedOn.ListIndex) <> -1 Then
If KeyCode = vbKeyF3 Then
  FrmEmpVacationSearch.index = 1
  Load FrmEmpVacationSearch

  FrmEmpVacationSearch.show vbModal
End If
End If
End Sub

Private Sub TxtOther_Change()
Smation
End Sub

Private Sub TxtPreSalary_Change()
Smation
End Sub

Private Sub TxtSalary_Change()
Smation
End Sub
Sub Smation()
'If Me.TxtModFlg.Text <> "R" Then
TxtTolaMostak.text = 0
TxtTotalCut.text = 0
If Ch(0).value = vbChecked Then
TxtTolaMostak.text = val(TxtTolaMostak.text) + val(TxtSalary.text)
End If
If Ch(1).value = vbChecked Then
TxtTolaMostak.text = val(TxtTolaMostak.text) + val(TxtIncrease.text)
End If
If Ch(2).value = vbChecked Then
TxtTolaMostak.text = val(TxtTolaMostak.text) + val(TxtSalaryVocation.text)
End If


If Ch(9).value = vbChecked Then
TxtDaySalVocation.text = val(val(TxtTotalDay.text)) + val(txtDaysCountPay.text)
'    If val(TxtDaySalary.text) <= 1 Then
'        TxtDaySalary.text = val(txtDaysCountPay.text)
'    Else
'        TxtDaySalary.text = val(TxtDaySalary.text) + val(txtDaysCountPay.text)
'    End If
Else
 '   TxtDaySalary = 1
End If


If Ch(3).value = vbChecked Then

TxtTolaMostak.text = val(TxtTolaMostak.text) + val(TxtSalEntitOther.text)
End If
If Ch(8).value = vbChecked Then
TxtTolaMostak.text = val(TxtTolaMostak.text) + val(TxtPreSalary.text)
End If
TxtTolaMostak.text = Round(val(TxtTolaMostak.text), 2)

If Ch(4).value = vbChecked Then
TxtTotalCut.text = val(TxtAdvance.text) + val(TxtTotalCut.text)
End If
If Ch(5).value = vbChecked Then
TxtTotalCut.text = val(TxtOther.text) + val(TxtTotalCut.text)
End If
If Ch(6).value = vbChecked Then
TxtTotalCut.text = val(TxtDecrease.text) + val(TxtTotalCut.text)
End If
If Ch(7).value = vbChecked Then
TxtTotalCut.text = val(TxtInsuranceValue.text) + val(TxtTotalCut.text)
End If
TxtTotalCut.text = Round(val(TxtTotalCut.text), 2)
'End If
End Sub
Private Sub TxtSalaryVocation_Change()
'TxtTolaMostak.text = val(Txtsalary.text) + val(Txtincrease.text) + val(TxtSalaryVocation.text) + val(TxtSalEntitOther.text)
Smation
End Sub

Private Sub TxtSalEntitOther_Change()
'TxtTolaMostak.text = val(Txtsalary.text) + val(Txtincrease.text) + val(TxtSalaryVocation.text) + val(TxtSalEntitOther.text)
Smation
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer
If Me.TxtModFlg.text <> "R" Then
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
        dateval
    End If
    End If
End Sub

Private Sub TxtToalAbsent_Change()
'TxtDuVocation.text = val(TxtContDay.text) - val(TxtToalAbsent.text) - val(TxtWithOutSala1.text)
End Sub

Private Sub TxtTolaMostak_Change()
NetTotal.text = val(TxtTolaMostak.text) - val(TxtTotalCut.text)
End Sub

Private Sub TxtTotalCut_Change()
NetTotal.text = val(TxtTolaMostak.text) - val(TxtTotalCut.text)
End Sub

Private Sub TxtTotalDay_Change()
If Me.TxtModFlg.text <> "R" Then
TxtDaySalVocation.text = val(TxtTotalDay.text)
End If
End Sub

Private Sub TxtValueTickt_Change()
If ChDelivery.value = vbChecked Then
Total.text = val(NetTotal.text) + val(TxtValueTickt.text)
Else
Total.text = val(NetTotal.text)
End If
End Sub

Private Sub TxtVSa_Change()
'txtToOutSal.text = (val(TxtVSa.text) + val(TxtMontOut.text) * 30 + val(TxtYaerOut.text) * 360)
'TxtDuVocation.text = val(TxtContDay.text) - val(TxtToalAbsent.text)
'DcboEmpName_Click (0)
End Sub

Private Sub TxtWithOutSala1_Change()
'DcboEmpName_Change
'TxtDuVocation.text = val(TxtContDay.text) - val(TxtToalAbsent.text) - val(TxtWithOutSala1.text)
End Sub

Private Sub TxtYaerOut_Change()
'txtToOutSal.text = val(TxtVSa.text) + val(TxtMontOut.text) * 30 + val(TxtYaerOut.text) * 360

'DcboEmpName_Click (0)
End Sub

Private Sub TxtYearAbs_Change()
'TxtAbsent.text = val(Me.TxtDayAbs.text) + val(Me.TxtMoAbs.text) * 30 + val(Me.TxtYearAbs.text) * 360
 
'TxtDuVocation.text = val(TxtContDay.text) - val(TxtToalAbsent.text)
'DcboEmpName_Click (0)
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
    Dim RsDev As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String

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
            rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    ''//
          If rs("chkGE").value = True Then
        chkGE.value = Checked
    Else
        chkGE.value = Unchecked
    End If
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
    
    stratDate.value = IIf(IsNull(rs("stratDate").value), Date, rs("stratDate").value)
    EndDate.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    '''/
    TxtDay2.text = IIf(IsNull(rs("TxtDay2").value), 0, rs("TxtDay2").value)
    TxtMonth2.text = IIf(IsNull(rs("TxtMonth2").value), 0, rs("TxtMonth2").value)
    TxtYear2.text = IIf(IsNull(rs("TxtYear2").value), 0, rs("TxtYear2").value)
    
    TxtDay3.text = IIf(IsNull(rs("TxtDay3").value), 0, rs("TxtDay3").value)
    TxtMonth3.text = IIf(IsNull(rs("TxtMonth3").value), 0, rs("TxtMonth3").value)
    TxtYear3.text = IIf(IsNull(rs("TxtYear3").value), 0, rs("TxtYear3").value)
    TxtAddDay.text = IIf(IsNull(rs("TxtAddDay").value), 0, rs("TxtAddDay").value)
    TxtDiscouDay.text = IIf(IsNull(rs("TxtDiscouDay").value), 0, rs("TxtDiscouDay").value)
    ''//////
    IDes.text = IIf(IsNull(rs("IDes").value), "0,0", rs("IDes").value)
    TxtNoVaction.text = IIf(IsNull(rs("TxtNoVaction").value), 0, rs("TxtNoVaction").value)
    LastBalanceMonth.text = IIf(IsNull(rs("LastBalanceMonth").value), 0, rs("LastBalanceMonth").value)
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    DateSta.value = IIf(IsNull(rs("DateSta").value), Date, rs("DateSta").value)
    Opretot.BoundText = IIf(IsNull(rs("OpretotID").value), 0, rs("OpretotID").value)
    
    ADDACC.BoundText = IIf(IsNull(rs("ADDACC").value), 0, rs("ADDACC").value)
    DISACC.BoundText = IIf(IsNull(rs("DISACC").value), 0, rs("DISACC").value)
    
    txtDaysCountPay = IIf(IsNull(rs("DaysCountPay").value), 0, rs("DaysCountPay").value)
    DCboUserName.BoundText = val(IIf(IsNull(rs("UserID").value), 0, rs("UserID").value))
    Me.DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcboJobsType.BoundText = IIf(IsNull(rs("JobID").value), "", rs("JobID").value)
    DcbDept.BoundText = IIf(IsNull(rs("DeptID").value), "", rs("DeptID").value)
    BignDate.value = IIf(IsNull(rs("BignDate").value), Date, rs("BignDate").value)
    LastVocatinDate.value = IIf(IsNull(rs("LastVocatinDate").value), Date, rs("LastVocatinDate").value)
    TxtContDay.text = IIf(IsNull(rs("ContDay").value), 0, rs("ContDay").value)
    TxtLastDayVoc.text = IIf(IsNull(rs("LastDayVoc").value), 0, rs("LastDayVoc").value)
    TxtDay.text = IIf(IsNull(rs("NoDay").value), 0, rs("NoDay").value)
    TxtMonth.text = IIf(IsNull(rs("NoMonth").value), 0, rs("NoMonth").value)
    TxtYear.text = IIf(IsNull(rs("NoYear").value), 0, rs("NoYear").value)
    TxtRemark.text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    TxtDaySalary.text = IIf(IsNull(rs("DaySalary").value), 0, rs("DaySalary").value)
    TxtSalary.text = IIf(IsNull(rs("Salary").value), 0, rs("Salary").value)
    txtPaymentRecommended.text = IIf(IsNull(rs("PaymentRecommended").value), 0, rs("PaymentRecommended").value)
   
    
    TxtDayIncrease.text = IIf(IsNull(rs("DayIncrease").value), 0, rs("DayIncrease").value)
    TxtIncrease.text = IIf(IsNull(rs("Increase").value), 0, rs("Increase").value)
    TxtDecrease.text = IIf(IsNull(rs("Decrease").value), 0, rs("Decrease").value)
    Me.dctype.BoundText = IIf(IsNull(rs("TypEndService").value), 0, rs("TypEndService").value)
    
    TxtDaySalVocation.text = IIf(IsNull(rs("DaySalVocation").value), 0, rs("DaySalVocation").value)
    TxtSalaryVocation.text = IIf(IsNull(rs("SalaryVocation").value), 0, rs("SalaryVocation").value)
    TxtDayEntitOther.text = IIf(IsNull(rs("DayEntitOther").value), 0, rs("DayEntitOther").value)
    TxtSalEntitOther.text = IIf(IsNull(rs("SalEntitOther").value), 0, rs("SalEntitOther").value)
    TxtOther.text = IIf(IsNull(rs("Other").value), 0, rs("Other").value)
    TxtAdvance.text = IIf(IsNull(rs("Advance").value), 0, rs("Advance").value)
    TxtValueTickt.text = IIf(IsNull(rs("ValueTickt").value), 0, rs("ValueTickt").value)
    Me.TxtTotalDay.text = IIf(IsNull(rs("TotalDay").value), 0, rs("TotalDay").value)
    ''///// 02 09 2015
    Me.TxtDuVocation.text = IIf(IsNull(rs("DuVocation").value), 0, rs("DuVocation").value)
    Me.TxtToalAbsent.text = IIf(IsNull(rs("ToalAbsent").value), 0, rs("ToalAbsent").value)
    Me.TxtYearAbs.text = IIf(IsNull(rs("YearAbs").value), 0, rs("YearAbs").value)
    Me.TxtMoAbs.text = IIf(IsNull(rs("MoAbs").value), 0, rs("MoAbs").value)
    Me.TxtDayAbs.text = IIf(IsNull(rs("DayAbs").value), 0, rs("DayAbs").value)
    Me.TxtVSa.text = IIf(IsNull(rs("DayOut").value), 0, rs("DayOut").value)
    Me.TxtMontOut.text = IIf(IsNull(rs("MontOut").value), 0, rs("MontOut").value)
    Me.TxtYaerOut.text = IIf(IsNull(rs("YaerOut").value), 0, rs("YaerOut").value)
    Me.Total.text = IIf(IsNull(rs("LastTotal").value), 0, rs("LastTotal").value)
    '''
    Me.TxtInsuranceValue.text = IIf(IsNull(rs("InsuranceValue").value), 0, rs("InsuranceValue").value)
    Me.TxtGetInsurance.text = IIf(IsNull(rs("GetInsurance").value), 0, rs("GetInsurance").value)
    Me.TxtNoMonth.text = IIf(IsNull(rs("NoMonth").value), 0, rs("NoMonth").value)
    CbBasedOn.ListIndex = IIf(IsNull(rs("BasedOn").value), -1, rs("BasedOn").value)
    TxtWithOutSala1.text = IIf(IsNull(rs("WithoutSala1").value), "", rs("WithoutSala1").value)
    Me.Txtorder.text = IIf(IsNull(rs("NoOrder").value), "", rs("NoOrder").value)
    TxtNoVacation.text = IIf(IsNull(rs("NoVacation").value), "", rs("NoVacation").value)
    TxtNewAbsent.text = IIf(IsNull(rs("NewAbsent").value), "", rs("NewAbsent").value)
    Me.TxtNoteID.text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)

Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)


 If IsNull(rs("Vact_Work").value) Or rs("Vact_Work").value = 0 Then
           Opt(0).value = True
           ElseIf rs("Vact_Work").value = 1 Then
           Opt(1).value = True
           ElseIf rs("Vact_Work").value = 2 Then
           Opt(2).value = True
           End If
    '''//////
  '   Me.TxtTolaMostak.text = IIf(IsNull(rs("TotalDue").value), 0, rs("TotalDue").value)
  '   Me.NetTotal.text = IIf(IsNull(rs("NetDue").value), 0, rs("NetDue").value)
  '   Me.TxtTotalCut.text = IIf(IsNull(rs("TotalCut").value), 0, rs("TotalCut").value)
  '   Me.Total.text = IIf(IsNull(rs("NetTotal").value), 0, rs("NetTotal").value)
    ''//
        If rs("ch0").value = True Then
             Ch(0).value = vbChecked
       Else
             Ch(0).value = vbUnchecked
       End If
          If rs("ch1").value = True Then
             Ch(1).value = vbChecked
       Else
             Ch(1).value = vbUnchecked
       End If
          If rs("ch2").value = True Then
             Ch(2).value = vbChecked
       Else
             Ch(2).value = vbUnchecked
       End If
          If rs("ch3").value = True Then
             Ch(3).value = vbChecked
       Else
             Ch(3).value = vbUnchecked
       End If
              If rs("ch4").value = True Then
             Ch(4).value = vbChecked
       Else
             Ch(4).value = vbUnchecked
       End If
       If rs("ch5").value = True Then
             Ch(5).value = vbChecked
       Else
             Ch(5).value = vbUnchecked
       End If
        If rs("ch6").value = True Then
             Ch(6).value = vbChecked
       Else
             Ch(6).value = vbUnchecked
       End If
       If rs("ch7").value = True Then
             Ch(7).value = vbChecked
       Else
             Ch(7).value = vbUnchecked
       End If
       If rs("ch8").value = 1 Then
             Ch(8).value = vbChecked
       Else
             Ch(8).value = vbUnchecked
       End If
       If rs("ch9").value = Null Then
            Ch(9).value = vbUnchecked
        Else
            If rs("ch9").value Then
                  Ch(9).value = vbChecked
            Else
                  Ch(9).value = vbUnchecked
            End If
        End If
       
       TxtPreSalary.text = IIf(IsNull(rs("PreSalary").value), 0, rs("PreSalary").value)
   '''''''''''''/
    
    If rs("Booked").value = True Then
             ChBooked.value = vbChecked
       Else
             ChBooked.value = vbUnchecked
    End If
    If rs("Delivery").value = True Then
             ChDelivery = vbChecked
        Else
             ChDelivery.value = vbUnchecked
    End If
    ''' aladein ADD
     If rs("Chekk").value = 0 Then
     Option1.value = vbChecked
     Else
     Option2.value = vbChecked
     End If
      
  '     If IsNull(rs("posted").value) Then
  '                                                 If SystemOptions.UserInterface = ArabicInterface Then
  '                                                  Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
  '                                                Else
  '                                                  Accredit.Caption = " send to Approval   "
  '                                             End If
  '                                             Accredit.Enabled = True
'  Else
'                                                   If SystemOptions.UserInterface = ArabicInterface Then
'                                                    Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
'                                                  Else
'                                                    Accredit.Caption = " sent to Approval   "
'                                               End If
'                                               Accredit.Enabled = False
'   End If
ShowComponent

StrSQL = " SELECT     dbo.TblVocationEntitlementsDet.VoEntID, dbo.TblVocationEntitlementsDet.Valu, dbo.TblVocationEntitlementsDet.TypeM, dbo.TblVocationEntitlementsDet.MofrdID, "
StrSQL = StrSQL & "                      dbo.mofrad.name , dbo.mofrad.NameE"
StrSQL = StrSQL & " FROM         dbo.TblVocationEntitlementsDet LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.mofrad ON dbo.TblVocationEntitlementsDet.MofrdID = dbo.mofrad.id"
StrSQL = StrSQL & "   Where (dbo.TblVocationEntitlementsDet.TypeM = 0) And (dbo.TblVocationEntitlementsDet.VoEntID = " & val(XPTxtID.text) & ")"
Set RsDev = New ADODB.Recordset
         VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.rows = 1
       RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsDev.RecordCount > 0) Then
        RsDev.MoveFirst
    
        With Me.Fg
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
 
                 .TextMatrix(i, .ColIndex("Serial")) = i
            
                .TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(RsDev("Valu").value), "", RsDev("Valu").value)
            
                .TextMatrix(i, .ColIndex("MofrdID")) = IIf(IsNull(RsDev("MofrdID").value), "", RsDev("MofrdID").value)
            
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                Else
                
                .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("nameE").value), "", RsDev("nameE").value)
               End If
                          
                RsDev.MoveNext
            Next i
 
        End With

    End If
    ''///
     VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
           VSFlexGrid1.rows = 1
    StrSQL = " SELECT     dbo.TblVocationEntitlementsDet.VoEntID, dbo.TblVocationEntitlementsDet.TypeM, dbo.TblVocationEntitlementsDet.MofrdID, "
StrSQL = StrSQL & "                       dbo.TblVocationEntitlementsDet.DeliverDate, dbo.TblVocationEntitlementsDet.ReciveDate, dbo.TblVocationEntitlementsDet.EmpID, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Namee , dbo.TblAssestes.AsName, dbo.TblAssestes.AsCode"
StrSQL = StrSQL & "  FROM         dbo.TblVocationEntitlementsDet LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblAssestes ON dbo.TblVocationEntitlementsDet.MofrdID = dbo.TblAssestes.AsID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.TblVocationEntitlementsDet.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "   Where (dbo.TblVocationEntitlementsDet.TypeM = 1) And (dbo.TblVocationEntitlementsDet.VoEntID = " & val(XPTxtID.text) & ")"
Set RsDev = New ADODB.Recordset
       RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If (RsDev.RecordCount > 0) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid1
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
 
               .TextMatrix(i, .ColIndex("Serial")) = i
            .TextMatrix(i, .ColIndex("AsCode")) = IIf(IsNull(RsDev("MofrdID").value), "", RsDev("MofrdID").value)
                .TextMatrix(i, .ColIndex("DeliverDate")) = IIf(IsNull(RsDev("DeliverDate").value), "", RsDev("DeliverDate").value)
                .TextMatrix(i, .ColIndex("ReciveDate")) = IIf(IsNull(RsDev("ReciveDate").value), "", RsDev("ReciveDate").value)
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(RsDev("EmpID").value), "", RsDev("EmpID").value)
                .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("AsName").value), "", RsDev("AsName").value)
                .TextMatrix(i, .ColIndex("MofrdID")) = IIf(IsNull(RsDev("MofrdID").value), "", RsDev("MofrdID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
                   Else
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
               
               End If
                          
                RsDev.MoveNext
            Next i
 
        End With

    End If
   ReLineGrid
   fillapprovData
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

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
        VSFlexGrid2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       VSFlexGrid2.TextMatrix(Num, VSFlexGrid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If VSFlexGrid2.TextMatrix(Num, VSFlexGrid2.ColIndex("Currcursor")) = "1" Then
   VSFlexGrid2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    VSFlexGrid2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        VSFlexGrid2.TextMatrix(Num, VSFlexGrid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            VSFlexGrid2.TextMatrix(Num, VSFlexGrid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             VSFlexGrid2.TextMatrix(Num, VSFlexGrid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            VSFlexGrid2.TextMatrix(Num, VSFlexGrid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            VSFlexGrid2.TextMatrix(Num, VSFlexGrid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            VSFlexGrid2.TextMatrix(Num, VSFlexGrid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          VSFlexGrid2.TextMatrix(Num, VSFlexGrid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If VSFlexGrid2.TextMatrix(Num, VSFlexGrid2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label10.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label10.Caption = "Approved"
                                 End If
                            Label10.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label10.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label10.Caption = "Currently required Approve"
                            End If
                 Label10.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 VSFlexGrid2.rows = 1
    End If
RsDetails.Close

End Function
Public Sub CreatLog_File_for_error(str As String)
    Dim StrLogFileName As String
    Dim IntFreeFile As Integer
    Dim ss As String

    StrLogFileName = App.path & "\employee_account_error.txt"

    If Dir(StrLogFileName) <> "" Then
        Kill StrLogFileName
    End If

    ss = "»Ì«‰ »«”„«¡ «·„ÊŸ›Ì‰ «·–Ì‰ ·œÌÂ„ „‘«ﬂ·  "
    ss = ss & vbCrLf & "Byte Informations Systems "
    ss = ss & vbCrLf & "BYTE "
    ss = ss & vbCrLf & "Create Date:- " & Now
    ss = ss & vbCrLf & str & vbCrLf
    IntFreeFile = FreeFile

    Open StrLogFileName For Output As #IntFreeFile
    Print #IntFreeFile, ss
    Close #IntFreeFile
End Sub

Function check_employee_accounts() As Boolean
    Dim Employee_account As String
    Dim error_string As String
    error_string = ""
    check_employee_accounts = True
    Dim i As Integer

    With Grid

        For i = .FixedRows To .rows - 2
                   If val(.TextMatrix(i, .ColIndex("BranchId"))) = 0 Then
                   error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡    ÕœÌœ «·›—⁄ «· «»⁄ ·Â"
        
                check_employee_accounts = False
                   End If
                   
                   
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code")

            If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«» –„ …"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«» –„ … ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
  Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
                    If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«» «·«ÃÊ— «·„” Õﬁ…"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«» «·«ÃÊ— «·„” Õﬁ… ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
            
  Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3")
                    If Employee_account = "" Or (Employee_account) = Null Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "   Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & vbCrLf & "·„ Ì „ «‰‘«¡ Õ”«»   «·„œ›Ê⁄«  «·„ﬁœ„…"
        
                check_employee_accounts = False
            End If
       
            If check_account_exist(Employee_account) = False Then
                error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & "    „ Õ–›  Õ”«»    «·„œ›Ê⁄«  «·„ﬁœ„… ÌœÊÌ« „‰ œ·Ì· «·Õ”«»«   " & vbCrLf
       
                check_employee_accounts = False
            End If
            
            
            '     If Val(.TextMatrix(i, .ColIndex("Emp_Salary"))) = 0 Then
            '     error_string = error_string + "  «·„ÊŸ› —ﬁ„ :" & .TextMatrix(i, .ColIndex("Emp_code")) & "  Ê«”„Â " & .TextMatrix(i, .ColIndex("Emp_Name")) & " ·„ Ì „  ÕœÌœ —« » «”«”Ì ·Â  " & vbCrLf
            '
            '    check_employee_accounts = False
            '
            '     End If
            If error_string <> "" Then
            CreatLog_File_for_error (error_string)
       End If
        Next i

    End With

    Dim X As Integer
    Dim StrLogFileName As String

    If error_string <> "" Then
        X = MsgBox("Â·  —Ìœ › Õ «·„·› ··„—«Ã⁄Â", vbCritical + vbYesNo, "ÌÊÃœ Œÿ√ ›Ì Õ”«»«  «·„ÊŸ›Ì‰  Ì„ﬂ‰ „—«Ã⁄ … ›Ì „·› «·«Œÿ«¡")

        If X = vbYes Then
            StrLogFileName = App.path & "\employee_account_error.txt"
            ShellExecute 0&, vbNullString, StrLogFileName, vbNullString, vbNullString, vbNormalFocus
        End If
    End If

End Function


Function GetComponentValuePerBranch(BramchId As Integer, componentname As String) As Double
    Dim SUM As Double
    SUM = 0
    Dim i As Integer

    With Grid

        For i = .FixedRows To .rows - 2
    
            If val(.TextMatrix(i, .ColIndex(componentname))) > 0 And val(.TextMatrix(i, .ColIndex("BranchId"))) = BramchId Then
                SUM = SUM + val(.TextMatrix(i, .ColIndex(componentname)))
            End If

        Next i

    End With

    GetComponentValuePerBranch = SUM
End Function

Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords



    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
        Msg = "«À»«  «” Õﬁ«ﬁ „ÊŸ› »”‰œ —ﬁ„" & XPTxtID & " ··„ÊŸ› " & DcboEmpName.text
    If check_employee_accounts = False Then
        Exit Function
    End If

 
        
BasicSalaryAccount = ""
 notes_id = general_noteid
                  
    For j = 1 To 40
        ColumnName = "Comp" & j

        If ViewComp(j) = True Then
                                  
            If CheckAccountToJE(Account_code(j)) = False Then
                Account_code(j) = SalaryAccount
            End If
        End If
    
    Next j
        
 
  
     my_branch = val(Dcbranch.BoundText)
 
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    '«·ÿ—› «·„œÌ‰ «·«÷«ﬁ« 
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim projectId As Integer
    
    BranchID = 1
    
    With Grid
BranchID = .TextMatrix(1, .ColIndex("BranchId"))
End With
    With Grid

        For j = 1 To 40

  '          If rsBranch.RecordCount > 0 Then
  '              rsBranch.MoveFirst
  '          End If
'
            ColumnName = "Comp" & j

            If ViewComp(j) = True And AddOrDiscount(j) = 0 Then '«·ŸÂÊ— Ê«÷«›… Ê·Ì” –„„ Ê·Ì” „ﬁœ„
                       If BasicSalaryAccount = "" Then
                                                                        BasicSalaryAccount = Account_code(j)
                                                 End If
                                                 
                If ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True Then
                                   
'                    For branch = 1 To rsBranch.RecordCount

'                        branchid = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                                
                        CValue = GetComponentValuePerBranch(BranchID, ColumnName)
                               If val(TxtDaySalary.text) = 0 Then
                               CValue = 0
                        
                                                 
                               End If
                               
                        If CValue > 0 Then
                                            
                    '    Debug.Print CValue & "  " & ColumnName
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 0, Msg & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & TxtDaySalary & " ÌÊ„  " & .TextMatrix(0, .ColIndex(ColumnName)), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If

                            line_no = line_no + 1
                        End If

 '                       rsBranch.MoveNext
'                    Next branch
                             
                End If
                             
            End If
    
        Next j
       
                                      
        For j = 1 To 40 ' Œ’Ê„« 

 '           If rsBranch.RecordCount > 0 Then
 '               rsBranch.MoveFirst
 '           End If

            ColumnName = "Comp" & j

            If ViewComp(j) = True And AddOrDiscount(j) = -1 Then
                If ZmamAccount(j) <> True And AdvPaymentdAccount(j) <> True Then
                                                                   
                    'For branch = 1 To rsBranch.RecordCount
                                 
 '                       branchid = IIf(IsNull(rsBranch("branch_id").value), 1, (rsBranch("branch_id").value))
                                
                        CValue = GetComponentValuePerBranch(BranchID, ColumnName)
                               If val(TxtDaySalary.text) = 0 Then
                               CValue = 0
                               End If
                               
                        If CValue > 0 Then
                   '      SystemOptions.ProjectEmployeeGV = True
 If SystemOptions.ProjectDiscountPolicy = 1 Then
                              
                            If ModAccounts.AddNewDev(LngDevID, line_no, Account_code1(j), CValue, 1, Msg & "   Œ’Ê„«  " & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & TxtDaySalary & " ÌÊ„  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If
                            
                            Else
                            
                                     If ModAccounts.AddNewDev(LngDevID, line_no, Account_code(j), CValue, 1, Msg & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & TxtDaySalary & " ÌÊ„  " & "   Œ’Ê„«  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If
                            
                            
                            
End If
                            line_no = line_no + 1
                        End If
                                    
 '                       rsBranch.MoveNext
 '                   Next branch
 '
                End If
            End If
    
        Next j

        For i = .FixedRows To .rows - 2
    
            If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) > 0 And val(val(TxtDaySalary.text)) <> 0 Then        '«·«ÃÊ— «·„” Õﬁ… œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & TxtDaySalary & " ÌÊ„  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
     
     
                 If val(.TextMatrix(i, .ColIndex("EmpTotalNet"))) < 0 And val(val(TxtDaySalary.text)) <> 0 Then         '«·«ÃÊ— «·„” Õﬁ… „œÌ‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, Abs(.TextMatrix(i, .ColIndex("EmpTotalNet"))), 0, Msg & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & TxtDaySalary & " ÌÊ„  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
            
            
            
            
            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True And ZmamAccount(j) = True Then
                     
                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                    StrAccountCode = Employee_account
                                                        
                    If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 And val(val(TxtDaySalary.text)) <> 0 Then
                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & " –„„ " & " —« » «·‘Â— «·Õ«·Ì »⁄œœ  " & TxtDaySalary & " ÌÊ„  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                            GoTo ErrTrap
                        End If

                        line_no = line_no + 1
                    End If
                 
                End If

            Next j
                 
            If val(.TextMatrix(i, .ColIndex("TotalAdvance"))) > 0 And val(val(TxtDaySalary.text)) <> 0 Then        '«·”·› œ«∆‰
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("TotalAdvance")), 1, Msg & "”œ«œ ”·› " & " —« » «·‘Â— «·Õ«·Ì    ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
 
 

                                                 
                                                 
'*********************************************************«Œ—Ì «÷«›…
CValue = val(TxtSalEntitOther)
Dim LINEACCOUNT As String


If CValue > 0 Then

If ADDACC.BoundText = "" Then
LINEACCOUNT = BasicSalaryAccount
Else
LINEACCOUNT = ADDACC.BoundText
End If


                       If ModAccounts.AddNewDev(LngDevID, line_no, LINEACCOUNT, CValue, 0, Msg & "   «Œ—Ì «÷«›… ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If


                   line_no = line_no + 1


                 Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, CValue, 1, Msg & "   «Œ—Ì «÷«›…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
End If


'*************************************************************************************************************************
'*********************************************************«Œ—Ì Œ’„
CValue = val(TxtOther)
If CValue > 0 Then


                 Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, CValue, 0, Msg & "   «Œ—Ì Œ’„   ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
If DISACC.BoundText = "" Then
LINEACCOUNT = BasicSalaryAccount
Else
LINEACCOUNT = DISACC.BoundText
End If
                
                
                
                                       If ModAccounts.AddNewDev(LngDevID, line_no, LINEACCOUNT, CValue, 1, Msg & "   «Œ—Ì Œ’„ ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                                GoTo ErrTrap
                            End If


                   line_no = line_no + 1


End If


'****************************************************************************”·› ﬂ«ﬂ·…
CValue = val(TxtAdvance)
             If CValue > 0 Then  '«·”·› œ«∆‰
             
                          Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") '«·«ÃÊ— «·„” Õﬁ…
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, CValue, 0, Msg & "   ”œ«œ ”·›     ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
                
                Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code") '–„Â
                StrAccountCode = Employee_account
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, CValue, 1, Msg & "”œ«œ ”·› ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            End If
            
 '****************************************************************************”·› ﬂ«ﬂ·…
'*******************************„œ›Ê⁄«  „ﬁ
            For j = 1 To 40
                ColumnName = "Comp" & j

                If ViewComp(j) = True And AdvPaymentdAccount(j) = True Then
                     
                    Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code3") 'œ›⁄«  „ﬁœ„…
                    StrAccountCode = Employee_account
                                 If AddOrDiscount(j) = 0 Then
                                                    If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 0, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                                            GoTo ErrTrap
                                                        End If
                                
                                                        line_no = line_no + 1
                                                    End If
                        
                        Else
                        
                                                                            If val(.TextMatrix(i, .ColIndex(ColumnName))) > 0 Then
                                                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex(ColumnName)), 1, Msg & "  „œ›Ê⁄«  „ﬁœ„…  ", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , val(.TextMatrix(i, .ColIndex("BranchId"))), , , , , , , , val(.TextMatrix(i, .ColIndex("Emp_ID")))) = False Then
                                                            GoTo ErrTrap
                                                        End If
                                
                                                        line_no = line_no + 1
                                                    End If

                        
                        
                        End If
                        
                 
                End If

            Next j
                 

            
'*******************************„œ›Ê⁄«  „ﬁ
 
        Next i

    End With
 SystemOptions.ProjectEmployeeGV = False

  If SystemOptions.ProjectEmployeeGV = True Then
'rs.Close
    Dim sql As String
    
    Dim Balance As Double
Dim mofradAccount As String
Dim mofradAccount1 As String
Dim Emp_id As Double
Dim Salary_account As String
 Dim Project_name As String
 Dim mofradname As String
  Dim AddOrDiscount1 As Integer
        sql = "SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.mofrad.Account_Code AS mofradAccount,  dbo.mofrad.Account_Code1 AS mofradAccount1, dbo.TblChangedComponentRegisterDetails.projectid,"
sql = sql & " dbo.Projects.Salary_account , dbo.Projects.Project_name, dbo.MOFRAD.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount"
sql = sql & " FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & "                       dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 0) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(" & SQLDate(NoteDate, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR(" & SQLDate(NoteDate, True) & "))"
sql = sql & " GROUP BY dbo.mofrad.Account_Code,dbo.mofrad.Account_Code1, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount"
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText 'stop

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("mofradAccount").value), "", rs("mofradAccount").value)
     mofradAccount1 = IIf(IsNull(rs("mofradAccount1").value), "", rs("mofradAccount1").value)
     
    'mofradAccount1
     
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), 0, rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & "", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    '
                     '            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , Notedate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchId) = False Then
                     '       GoTo ErrTrap
                     '   End If
        
        
    
                             If SystemOptions.ProjectDiscountPolicy = 1 Then
                             
                                        If mofradAccount1 <> "" Then
                                        Salary_account = mofradAccount1
                                        End If
                            
                             
                             End If
                             
                                line_no = line_no + 1
                                                             If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
                        
                        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
        
            line_no = line_no + 1
        
        
         
                        
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
    
    
    
    
    
    
    
    
  
    
'«·„‘«—Ì⁄ Ê·ﬂ‰ –„„
 Dim empAccount_Codezmam As String
 Dim emp_Name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.ZmamAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(  " & SQLDate(NoteDate, True) & " )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(NoteDate, True) & " ))"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code"
 
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText '0000000

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close

    
    
   ' Õ„Ì· «·„’—Ê›«  ⁄·Ï «·„‘«—Ì⁄
    
       sql = "SELECT      SUM(ROUND(dbo.EmpSalaryComponent.[Value] * dbo.opr_employee_details.[interval] / 30, 2)) AS Total, dbo.mofrad.Account_Code, "
sql = sql & " dbo.mofrad.AddOrDiscount, dbo.EmpSalaryComponent.EntIncresDataM, dbo.projects.Salary_account, 2006 + dbo.opr_Employee.Years AS [year],"
sql = sql & " dbo.opr_Employee.Months, SUM(dbo.opr_employee_details.[interval]) AS Intervals, dbo.opr_employee_details.ProjectID, dbo.mofrdat.mofrad_name,"
sql = sql & " dbo.Projects.Project_name , dbo.TblEmployee.BranchId"
sql = sql & " FROM         dbo.opr_employee_details INNER JOIN"
sql = sql & " dbo.projects ON dbo.opr_employee_details.ProjectID = dbo.projects.id INNER JOIN"
sql = sql & " dbo.opr_Employee ON dbo.opr_employee_details.pk_id = dbo.opr_Employee.id INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.opr_employee_details.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.EmpSalaryComponent ON dbo.opr_employee_details.Emp_id = dbo.EmpSalaryComponent.emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad INNER JOIN"
sql = sql & " dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type ON dbo.EmpSalaryComponent.AccountCode = dbo.mofrdat.mofrad_code"
sql = sql & " GROUP BY dbo.mofrad.Account_Code, dbo.EmpSalaryComponent.EntIncresDataM, dbo.projects.Salary_account, 2006 + dbo.opr_Employee.Years, dbo.opr_Employee.Months,"
sql = sql & " dbo.MOFRAD.AddOrDiscount , dbo.opr_employee_details.ProjectID, dbo.mofrdat.mofrad_name, dbo.Projects.Project_name, dbo.TblEmployee.BranchId"
sql = sql & " HAVING      (dbo.EmpSalaryComponent.EntIncresDataM IS NULL  OR"
sql = sql & "  dbo.EmpSalaryComponent.EntIncresDataM >= " & SQLDate(NoteDate, True) & " )"

sql = sql & "   AND (dbo.opr_Employee.Months = " & CmbMonth.ListIndex & ") AND (2006 + dbo.opr_Employee.Years = " & val(CboYear.text) & ")"


sql = sql & " ORDER BY dbo.opr_employee_details.ProjectID"

 
   
  
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     mofradAccount = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("mofrad_name").value), "", rs("mofrad_name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
             projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
             If mofradAccount <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, mofradAccount, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    
    
    
    
    
    
    
    
'«·„‘«—Ì⁄ Ê·ﬂ‰ œ›⁄«  „ﬁœ„…
 'Dim empAccount_Codezmam As String
 'Dim emp_name As String
            sql = " SELECT     SUM(dbo.TblChangedComponentRegisterDetails.[value]) AS Balance, dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account,"
sql = sql & " dbo.projects.Project_name, dbo.mofrad.name, dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code,"
sql = sql & " dbo.TblEmployee.emp_name , dbo.TblEmployee.Account_Code3"
sql = sql & "  FROM         dbo.TblChangedComponentRegister INNER JOIN"
sql = sql & " dbo.TblChangedComponentRegisterDetails ON"
sql = sql & " dbo.TblChangedComponentRegister.ChangedComponentid = dbo.TblChangedComponentRegisterDetails.ChangedComponentid INNER JOIN"
sql = sql & " dbo.TblEmployee ON dbo.TblChangedComponentRegisterDetails.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & " dbo.mofrad ON dbo.TblChangedComponentRegister.ComponentID = dbo.mofrad.id LEFT OUTER JOIN"
sql = sql & " dbo.projects ON dbo.TblChangedComponentRegisterDetails.projectid = dbo.projects.id"
sql = sql & " WHERE     (dbo.mofrad.AdvPaymentdAccount = 1) AND (MONTH(dbo.TblChangedComponentRegister.RecordDate) = MONTH(   " & SQLDate(NoteDate, True) & "  )) AND"
sql = sql & " (YEAR(dbo.TblChangedComponentRegister.RecordDate) = YEAR( " & SQLDate(NoteDate, True) & "  ))"
sql = sql & " GROUP BY dbo.TblChangedComponentRegisterDetails.projectid, dbo.projects.Salary_account, dbo.projects.Project_name, dbo.mofrad.name,"
sql = sql & " dbo.TblChangedComponentRegister.BranchId, dbo.mofrad.AddOrDiscount, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
sql = sql & " dbo.TblEmployee.Account_Code3"
 
 
    
  
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
    For i = 1 To rs.RecordCount
     empAccount_Codezmam = IIf(IsNull(rs("Account_Code3").value), "", rs("Account_Code3").value)
     Salary_account = IIf(IsNull(rs("Salary_account").value), "", rs("Salary_account").value)
     Balance = IIf(IsNull(rs("Balance").value), 0, rs("Balance").value)
     Project_name = IIf(IsNull(rs("Project_name").value), "", rs("Project_name").value)
     mofradname = IIf(IsNull(rs("name").value), "", rs("name").value)
     BranchID = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     AddOrDiscount1 = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
     emp_Name = IIf(IsNull(rs("emp_name").value), "", rs("emp_name").value)
     projectId = IIf(IsNull(rs("projectid").value), 0, rs("projectid").value)
     
             If empAccount_Codezmam <> "" And Salary_account <> "" And Balance > 0 Then
                   
                  If AddOrDiscount1 = 0 Then '«÷«›Ì
                   
                   If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                             
                    Else ' Œ’„
                    
                                 If ModAccounts.AddNewDev(LngDevID, line_no, empAccount_Codezmam, Balance, 0, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                        
                            If ModAccounts.AddNewDev(LngDevID, line_no, Salary_account, Balance, 1, Msg & mofradname & "  " & "··„‘—Ê⁄   " & Project_name & " ·  " & emp_Name, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , projectId, , , , , , , BranchID) = False Then
                            GoTo ErrTrap
                        End If
        
                        line_no = line_no + 1
                    
                    End If
                    
                             
                             
             End If
     rs.MoveNext
     Next i
    End If

    rs.Close
    

End If


'«· √„Ì‰« 


'    rs.Close
    
       
       sql = " "

'
 

' project gv

'    Create_dev2 = True
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
'    Create_dev2 = False
  
'********************************************************************
  End Function


Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "«À»«  «” Õﬁ«ﬁ „ÊŸ› »”‰œ —ﬁ„" & XPTxtID & " ··„ÊŸ› " & DcboEmpName.text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim sql As String
tablename = "TblVocationEntitlements"
Filedname = "ID"
NoteSerial1 = val(XPTxtID)
Notevalue = 0

 notytype = 8065
Notevalue = val(Total)
 

 BranchID = val(Dcbranch.BoundText)
 If chkGE.value = vbChecked Then
 NoteDate = (XPDtbTrans.value)
 Else
 NoteDate = (DateSta.value)
 End If
 


 
If Notevalue > 0 Then
                              '  If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TxtNoteID.text = NoteID
                                                     TxtNoteSerial.text = NoteSerial
                                   '  Else
                                         '        If TxtNoteID.text = "" Or TxtNoteSerial.text = "" Then
                                         '   CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                         '                        TxtNoteID.text = NoteID
                                         '                       TxtNoteSerial.text = NoteSerial
                                         '          Else
                                         '                        sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                         '                       sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                         '                          sql = sql & " where NoteID=" & val(TxtNoteID.text)
                                         '                          Cn.Execute sql
                                         '
                                         '        End If
                                       
                             '   End If

CREATE_VOUCHER_GE val(TxtNoteID.text), BranchID, user_id, NoteDate
rs.Resync adAffectCurrent
 

     End If

End Function

Private Sub GetAdvanceValuesSalary(IntMonth As Integer, _
                             IntYear As Integer)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim LngFindRow As Long
    On Error GoTo hErr
    StrSQL = "Select Emp_ID,Sum(TotalAdvance)as CCC From ( SELECT QryAllEmpAdvance.Emp_ID,QryA" & "llEmpAdvance.TotalAdvance FROM   dbo.QryAllEmpAdvance(" & IntMonth & "," & IntYear & ") QryAllEmpAdvance )" & "Xtable Group By Emp_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    With Me.Grid1
        rs.MoveFirst
        .cell(flexcpText, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance")) = 0

        For i = 1 To rs.RecordCount
            LngFindRow = .FindRow(rs("Emp_ID").value, .FixedRows, .ColIndex("Emp_ID"), False, True)

            If LngFindRow <> -1 Then
                If Not (IsNull(rs("CCC").value)) Then
                    .TextMatrix(LngFindRow, .ColIndex("TotalAdvance")) = Round(rs("CCC").value, 0)
                End If
            End If

            rs.MoveNext
        Next i

    End With

hErr:
    'Stop
End Sub

Private Sub GetAdvanceValues(IntMonth As Integer, _
                             IntYear As Integer)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim LngFindRow As Long
    On Error GoTo hErr
    StrSQL = "Select Emp_ID,Sum(TotalAdvance)as CCC From ( SELECT QryAllEmpAdvance.Emp_ID,QryA" & "llEmpAdvance.TotalAdvance FROM   dbo.QryAllEmpAdvance(" & IntMonth & "," & IntYear & ") QryAllEmpAdvance )" & "Xtable Group By Emp_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    With Me.Grid
        rs.MoveFirst
        .cell(flexcpText, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance")) = 0

        For i = 1 To rs.RecordCount
            LngFindRow = .FindRow(rs("Emp_ID").value, .FixedRows, .ColIndex("Emp_ID"), False, True)

            If LngFindRow <> -1 Then
                If Not (IsNull(rs("CCC").value)) Then
                    .TextMatrix(LngFindRow, .ColIndex("TotalAdvance")) = Round(rs("CCC").value, 0)
                End If
            End If

            rs.MoveNext
        Next i

    End With

hErr:
    'Stop
End Sub
Private Sub CalculateNetsSalary()
    Dim i As Integer
    Dim SngHourPrice As Single
    Dim SngOverTimePrice As Single

    Dim NetTotal As Single
    Dim SngTemp As Single
    Dim TotalAddtion As Double
    Dim TotalDiscount As Double
    Dim ColumnName As String
    Dim SngTotal As Double
    Dim j As Integer
    'On Error GoTo ErrTrap
    On Error Resume Next

    With Me.Grid1

        If .FixedRows = .rows Then Exit Sub

        For i = .FixedRows To .rows - 1

            TotalAddtion = 0
            TotalDiscount = 0

            For j = 1 To 40
                ColumnName = "Comp" & j

                If AddOrDiscount(j) = 0 Then
                    TotalAddtion = TotalAddtion + val(.TextMatrix(i, .ColIndex(ColumnName)))
                Else
                    TotalDiscount = TotalDiscount + val(.TextMatrix(i, .ColIndex(ColumnName)))
                End If

            Next j
        
            .TextMatrix(i, .ColIndex("total1")) = val(.TextMatrix(i, .ColIndex("Mokafea"))) + TotalAddtion
            .TextMatrix(i, .ColIndex("total2")) = val(.TextMatrix(i, .ColIndex("TotalAdvance"))) + val(.TextMatrix(i, .ColIndex("TotalDiscount"))) + TotalDiscount
            .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2")))

            If i Mod 2 = 0 Then
                .cell(flexcpBackColor, i, 1, i, 41) = &HE0E0E0
     
            End If
        
        Next i
    
    End With

    Exit Sub
ErrTrap:
    'Resume
End Sub
Private Sub CalculateNets()
    Dim i As Integer
    Dim SngHourPrice As Single
    Dim SngOverTimePrice As Single

    Dim NetTotal As Single
    Dim SngTemp As Single
    Dim TotalAddtion As Double
    Dim TotalDiscount As Double
    Dim ColumnName As String
    Dim SngTotal As Double
    Dim j As Integer
    'On Error GoTo ErrTrap
    On Error Resume Next

    With Me.Grid

        If .FixedRows = .rows Then Exit Sub

        For i = .FixedRows To .rows - 1

            TotalAddtion = 0
            TotalDiscount = 0

            For j = 1 To 40
                ColumnName = "Comp" & j

                If AddOrDiscount(j) = 0 Then
                    TotalAddtion = TotalAddtion + val(.TextMatrix(i, .ColIndex(ColumnName)))
                Else
                    TotalDiscount = TotalDiscount + val(.TextMatrix(i, .ColIndex(ColumnName)))
                End If

            Next j
        
            .TextMatrix(i, .ColIndex("total1")) = val(.TextMatrix(i, .ColIndex("Mokafea"))) + TotalAddtion
            .TextMatrix(i, .ColIndex("total2")) = val(.TextMatrix(i, .ColIndex("TotalAdvance"))) + val(.TextMatrix(i, .ColIndex("TotalDiscount"))) + TotalDiscount
            .TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2")))

            If i Mod 2 = 0 Then
                .cell(flexcpBackColor, i, 1, i, 41) = &HE0E0E0
     
            End If
        
        Next i
    
    End With

    Exit Sub
ErrTrap:
    'Resume
End Sub
Public Sub FillGridWithData()

    Dim i As Integer
    Dim j As Integer
    Dim countFlag As Integer
    Dim AllwIntro As Double
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
    Dim ColumnName As String
    Dim TotalAddtion As Double
    Dim TotalDiscount As Double
    Dim DaysInMonth22 As Double
    Dim CountDays22 As Double
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

countFlag = 1
 

    IntYear = year(DateSta.value)
    IntMonth = Month(DateSta.value)

      Grid.Clear flexClearScrollable, flexClearEverything
              Grid.rows = 1
              
        Dim ID As String
 
    My_SQL = " Select Emp_Namee, lastHolidaydate,BignDateWork,  fullcode,groupid,  BranchId,Emp_ID,Emp_Code,Emp_Name,DepartmentID,project_id ,cost_center_id,IsNUll(Emp_Salary,0)as Emp_Salary,IsNUll(Emp_Salary_sakn,0)as Emp_Salary_sakn,IsNUll(Emp_Salary_bus,0)as Emp_Salary_bus,IsNUll(Emp_Salary_food,0)as Emp_Salary_food,IsNUll(Emp_Salary_others,0)as Emp_Salary_others,IsNUll(Emp_Salary_mob,0)as Emp_Salary_mob,IsNUll(Emp_Salary_mang,0)as Emp_Salary_mang,  IsNUll( TotalDiscount,0)as TotalDiscount,IsNUll(TotalMokafea, 0) As TotalMokafea,(IsNUll(Emp_Salary,0)+IsNUll( TotalMokafea,0))-(IsNUll(TotalDiscount,0)) as EmpTotalNet ,JobTypeName, JobTypeNamee,branch_name,branch_namee,projectFullcode,Project_name,Project_nameE" & CHR(13)
  My_SQL = My_SQL + "  From (" & CHR(13)

  My_SQL = My_SQL + "  SELECT     TOP 100 PERCENT  dbo.TblEmployee.Emp_Namee , dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.BranchId, dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.cost_center_id, SUM(QryAllDiscountWithMkafea.TotalDiscount) AS TotalDiscount, SUM(QryAllDiscountWithMkafea.Mokafea) AS TotalMokafea," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee," & CHR(13)
  My_SQL = My_SQL + "                       dbo.projects.Fullcode AS projectFullcode, dbo.projects.Project_name, dbo.projects.Project_nameE" & CHR(13)
  My_SQL = My_SQL + " FROM         dbo.TblEmpJobsTypes INNER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.QryAllDiscountWithMkafea(" & IntMonth & ", " & IntYear & ") QryAllDiscountWithMkafea ON dbo.TblEmployee.Emp_ID = QryAllDiscountWithMkafea.Emp_ID" & CHR(13)

 
        My_SQL = My_SQL + " and dbo.TblEmployee.BignDateWork<" & SQLDate(DateSta.value, True)
                If DcboEmpName.text <> "" Then
            My_SQL = My_SQL + " Where  dbo.TblEmployee.Emp_id=" & val(DcboEmpName.BoundText) ' & "'"
        End If

 'DcboEmpName
 My_SQL = My_SQL + "  GROUP BY dbo.TblEmployee.Emp_Namee , dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID, dbo.TblEmployee.BranchId, " & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.cost_center_id, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.projects.Fullcode, dbo.projects.Project_name," & CHR(13)
My_SQL = My_SQL + "                      dbo.Projects.Project_nameE" & CHR(13)
My_SQL = My_SQL + " ORDER BY dbo.TblEmployee.Fullcode" & CHR(13)

My_SQL = My_SQL + "  )XTable"


    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst
Dim CountDays As Double
 
Dim MonthDayNo  As Double

MonthDayNo = daysInMonth(DateSta.value)

If MonthDayNo = 28 Then
MonthDayNo = 30
ElseIf MonthDayNo = 31 Then
MonthDayNo = 30
End If

            For i = 1 To .rows - 1
         countFlag = 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
            .TextMatrix(i, .ColIndex("BignDateWork")) = IIf(IsNull(rs.Fields("BignDateWork").value), "", rs.Fields("BignDateWork").value)
            .TextMatrix(i, .ColIndex("lastHolidaydate")) = IIf(IsNull(rs.Fields("lastHolidaydate").value), "", rs.Fields("lastHolidaydate").value)

           
           CountDays = day(DateSta.value)
           
           If MonthDayNo <= CountDays Then
CountDays = 30
 
End If

MonthDayNo = 30

   CountDays22 = day(DateSta.value)
   'Abs(DateDiff("D", MonthLastDay(DateSta.value), DateSta.value))
         '  CountDays22 = CountDays22 + 1
           DaysInMonth22 = daysInMonth(DateSta.value)
           
            .TextMatrix(i, .ColIndex("CountDays")) = CountDays
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), 1, rs.Fields("BranchId").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("fullcode").value), "", rs.Fields("fullcode").value)
                .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs.Fields("cost_center_id").value), "", rs.Fields("cost_center_id").value)
     
                
                      If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs.Fields("JobTypeName").value), "", rs.Fields("JobTypeName").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
           Else
           .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs.Fields("JobTypeNamee").value), "", rs.Fields("JobTypeNamee").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value)
           End If
                TotalAddtion = 0
                TotalDiscount = 0

                For j = 1 To 40
                    ColumnName = "Comp" & j

                    If ViewComp(j) = True Then
                    AllwIntro = GetValueAllwIntro(Month(DateSta.value), year(DateSta.value), val(DcboEmpName.BoundText), j)
                    If AllwIntro <= 0 Then
                        If FixedOrChanged(j) = 0 Then
                            .TextMatrix(i, .ColIndex(ColumnName)) = GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(i, .ColIndex("Emp_ID"))), CStr(j), , DateSta.value)
                                           
                                           If countFlag = 1 Then
                                           If showMofradAll(j) = False Then
                                            If culc30orRminder(j) = 0 Then
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * CountDays, 2)
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / DaysInMonth22 * CountDays22, 2)
                                          End If
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))), 2)
                                          End If
                                           End If
                                           
                        Else
                            .TextMatrix(i, .ColIndex(ColumnName)) = GetEmployeeChangedSalary(val(.TextMatrix(i, .ColIndex("Emp_ID"))), j, val(CboYear.text), CmbMonth.ListIndex + 1)
                           ' .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))), 2)
                          
                        End If
                       Else
                       .TextMatrix(i, .ColIndex(ColumnName)) = AllwIntro
                       End If
                    End If
    
                Next j
    
                 '         .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Round(rs.Fields("TotalDiscount").value, Decimal_Places))
             
                '.TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Round(rs.Fields("TotalMokafea").value, Decimal_Places))
              
                rs.MoveNext
            
            Next

            rs.Close
        End If

        GetAdvanceValues IntMonth, IntYear
        ' GetWorkHours
        CalculateNets
        .rows = .rows + 1

        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        Else
            .TextMatrix(.rows - 1, .ColIndex("Ser")) = "Total"
        End If

        .IsSubtotal(.rows - 1) = True
        Dim SngTotal As Single

        For j = 1 To 40
            ColumnName = "Comp" & j
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex(ColumnName), .rows - 1, .ColIndex(ColumnName))
            .TextMatrix(.rows - 1, .ColIndex(ColumnName)) = SngTotal
     
        Next j
 
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .rows, .ColIndex("total1"))
        .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
        TxtSalary = SngTotal
        
        If val(TxtDaySalary.text) <> 0 Then
        
        
        
       Else
            TxtSalary = 0
       End If
       
       If Ch(9).value = vbChecked Then
            
           ' TxtSalaryVocation = SngTotal * val(txtDaysCountPay.text)
            TxtDaySalVocation_Change
        Else
            
        End If
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .rows, .ColIndex("total2"))
        .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
         TxtDecrease.text = SngTotal
         If val(TxtDecrease.text) > 0 Then
        ' Ch(6).Enabled = False
         Ch(6).value = vbChecked
         Else
         Ch(6).Enabled = True
         Ch(6).value = vbUnchecked
         End If
         Smation
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows, .ColIndex("EmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
'

        .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
        .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
        .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
        .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With
 

'rs.Close
Set rs = Nothing

'    Coloring
ErrTrap:

End Sub
Public Function MonthLastDay(ByVal dCurrDate As Date) As Date
    Dim dFirstDayNextMonth As Date
    MonthLastDay = Empty
    dCurrDate = Format(dCurrDate, "DD/MM/YYYY")
    dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
    MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
    Exit Function
End Function
Public Sub FillGridWithDataSalary(Optional NoRow As Integer, Optional RecDate As Date)
    Dim i As Integer
    Dim j As Integer
    Dim countFlag As Integer
    Dim AllwIntro As Double
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
    Dim ColumnName As String
    Dim TotalAddtion As Double
    Dim TotalDiscount As Double
    Dim DaysInMonth22 As Double
    Dim CountDays22 As Double
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
If val(val(DcboEmpName.BoundText)) = 0 Or NoRow = 0 Then Exit Sub
countFlag = 1
 
    DataSalary.value = MonthLastDay(RecDate)
    IntYear = year(DataSalary.value)
    IntMonth = Month(DataSalary.value)

 
        Dim ID As String
 
 
    My_SQL = " Select  lastHolidaydate,BignDateWork,  fullcode,groupid,  BranchId,Emp_ID,Emp_Code,Emp_Name,DepartmentID,project_id ,cost_center_id,IsNUll(Emp_Salary,0)as Emp_Salary,IsNUll(Emp_Salary_sakn,0)as Emp_Salary_sakn,IsNUll(Emp_Salary_bus,0)as Emp_Salary_bus,IsNUll(Emp_Salary_food,0)as Emp_Salary_food,IsNUll(Emp_Salary_others,0)as Emp_Salary_others,IsNUll(Emp_Salary_mob,0)as Emp_Salary_mob,IsNUll(Emp_Salary_mang,0)as Emp_Salary_mang,  IsNUll( TotalDiscount,0)as TotalDiscount,IsNUll(TotalMokafea, 0) As TotalMokafea,(IsNUll(Emp_Salary,0)+IsNUll( TotalMokafea,0))-(IsNUll(TotalDiscount,0)) as EmpTotalNet ,JobTypeName, JobTypeNamee,branch_name,branch_namee,projectFullcode,Project_name,Project_nameE ,Emp_Namee" & CHR(13)
  My_SQL = My_SQL + "  From (" & CHR(13)

  My_SQL = My_SQL + "  SELECT     TOP 100 PERCENT dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.BranchId, dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee.cost_center_id, SUM(QryAllDiscountWithMkafea.TotalDiscount) AS TotalDiscount, SUM(QryAllDiscountWithMkafea.Mokafea) AS TotalMokafea," & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee," & CHR(13)
  My_SQL = My_SQL + "                       dbo.projects.Fullcode AS projectFullcode, dbo.projects.Project_name, dbo.projects.Project_nameE ,dbo.TblEmployee.Emp_Namee" & CHR(13)
  My_SQL = My_SQL + " FROM         dbo.TblEmpJobsTypes INNER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblEmployee ON dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.projects ON dbo.TblEmployee.project_id = dbo.projects.id LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN" & CHR(13)
  My_SQL = My_SQL + "                       dbo.QryAllDiscountWithMkafea(" & IntMonth & ", " & IntYear & ") QryAllDiscountWithMkafea ON dbo.TblEmployee.Emp_ID = QryAllDiscountWithMkafea.Emp_ID" & CHR(13)

 
        My_SQL = My_SQL + " and dbo.TblEmployee.BignDateWork<" & SQLDate(DataSalary.value, True)
                If DcboEmpName.text <> "" Then
            My_SQL = My_SQL + " Where  dbo.TblEmployee.Emp_id=" & val(DcboEmpName.BoundText) ' & "'"
        End If

 'DcboEmpName
 My_SQL = My_SQL + "  GROUP BY dbo.TblEmployee.lastHolidaydate, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.Fullcode, dbo.TblEmployee.GroupID, dbo.TblEmployee.BranchId, " & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmployee.cost_center_id, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.DepartmentID, dbo.TblEmployee.project_id, dbo.TblEmpJobsTypes.JobTypeName," & CHR(13)
My_SQL = My_SQL + "                      dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.projects.Fullcode, dbo.projects.Project_name," & CHR(13)
My_SQL = My_SQL + "                      dbo.Projects.Project_nameE ,dbo.TblEmployee.Emp_Namee " & CHR(13)
My_SQL = My_SQL + " ORDER BY dbo.TblEmployee.Fullcode" & CHR(13)

My_SQL = My_SQL + "  )XTable"
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    With Me.Grid1
        If rs.RecordCount > 0 Then
            rs.MoveFirst
Dim CountDays As Double
 
Dim MonthDayNo  As Double

MonthDayNo = daysInMonth(DataSalary.value)

If MonthDayNo = 28 Then
MonthDayNo = 30
ElseIf MonthDayNo = 31 Then
MonthDayNo = 30
End If

            For i = NoRow To NoRow
         countFlag = 1
                .TextMatrix(i, .ColIndex("Ser")) = i
           CountDays = day(DataSalary.value)
           
           If MonthDayNo <= CountDays Then
CountDays = 30
 
End If

MonthDayNo = 30

   CountDays22 = day(DataSalary.value)
           DaysInMonth22 = daysInMonth(DataSalary.value)
              '  .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("ID").value), "", rs.Fields("ID").value)
                .TextMatrix(i, .ColIndex("payed")) = 1
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), 1, rs.Fields("BranchId").value)
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("fullcode").value), "", rs.Fields("fullcode").value)
                .TextMatrix(i, .ColIndex("cost_center_id")) = IIf(IsNull(rs.Fields("cost_center_id").value), "", rs.Fields("cost_center_id").value)
                .TextMatrix(i, .ColIndex("RecordDate")) = DataSalary.value
         If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
           Else
          ' .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs.Fields("JobTypeNamee").value), "", rs.Fields("JobTypeNamee").value)
           .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value)
         End If
                TotalAddtion = 0
                TotalDiscount = 0

                For j = 1 To 40
                    ColumnName = "Comp" & j

                    If ViewComp(j) = True Then
                    AllwIntro = GetValueAllwIntro(Month(DataSalary.value), year(DataSalary.value), val(DcboEmpName.BoundText), j)
                    If AllwIntro <= 0 Then
                        If FixedOrChanged(j) = 0 Then
                            .TextMatrix(i, .ColIndex(ColumnName)) = GetEmployeeSalaryAccordingToComponent(val(.TextMatrix(i, .ColIndex("Emp_ID"))), CStr(j), , DataSalary.value)
                                           
                                           If countFlag = 1 Then
                                           If showMofradAll(j) = False Then
                                            If culc30orRminder(j) = 0 Then
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / MonthDayNo * CountDays, 2)
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))) / DaysInMonth22 * CountDays22, 2)
                                          End If
                                          Else
                                          .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))), 2)
                                          End If
                                           End If
                                           
                        Else
                            .TextMatrix(i, .ColIndex(ColumnName)) = GetEmployeeChangedSalary(val(.TextMatrix(i, .ColIndex("Emp_ID"))), j, val(CboYear.text), CmbMonth.ListIndex + 1)
                           ' .TextMatrix(i, .ColIndex(ColumnName)) = Round(val(.TextMatrix(i, .ColIndex(ColumnName))), 2)
                          
                        End If
                       Else
                       .TextMatrix(i, .ColIndex(ColumnName)) = AllwIntro
                       End If
                    End If
    
                Next j

                rs.MoveNext
            
            Next

            rs.Close
        End If
        GetAdvanceValuesSalary IntMonth, IntYear
        CalculateNetsSalary
RelinSalaryPayed
    End With
Set rs = Nothing
ErrTrap:
End Sub
Function GetValueAllwIntro(Optional MothID As Integer, Optional YerID As Integer, Optional EmpID As Double, Optional MofrdID As Integer) As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     MordValue / ISNULL(TypeMofrd, 1) AS Valu"
sql = sql & " From dbo.TblComponentYearDet"
sql = sql & " WHERE       (EmpID = " & EmpID & ") AND (MofrdID = " & MofrdID & ") and "
sql = sql & "               ((month(RecDate1) =" & MothID & " and Year(RecDate1) =" & YerID & ") or    ((month(RecDate2) =" & MothID & " and Year(RecDate2) =" & YerID & ")))"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetValueAllwIntro = IIf(IsNull(Rs3("Valu").value), 0, Rs3("Valu").value)
Else
GetValueAllwIntro = 0
End If
End Function
Private Sub ShowComponent()
    On Error Resume Next

If DcboEmpName.BoundText = "" Then Exit Sub
'firstrun = False
     If getTitlesName = True Then
   End If
    DoEvents
    FillGridWithData
 
    Dim i As Integer
        With Grid
For i = 1 To 40

                 If val((.TextMatrix(.rows - 1, .ColIndex("Comp" & i & "")))) = 0 Then
                   .ColHidden(.ColIndex("Comp" & i)) = True
                End If


                If val((.TextMatrix(.rows - 1, .ColIndex("sgn")))) = 0 Then
                  .ColHidden(.ColIndex("sgn")) = True
                End If
               If val((.TextMatrix(.rows - 1, .ColIndex("TotalAdvance")))) = 0 Then
                  .ColHidden(.ColIndex("TotalAdvance")) = True
                End If
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("TotalDiscount")))) = 0 Then
                  .ColHidden(.ColIndex("TotalDiscount")) = True
                  Else
                '  TxtDecrease.Text = val((.TextMatrix(.Rows - 1, .ColIndex("TotalDiscount"))))
                End If
                
                          If val((.TextMatrix(.rows - 1, .ColIndex("Mokafea")))) = 0 Then
                  .ColHidden(.ColIndex("Mokafea")) = True
                End If
Next i
End With
End Sub

Function getTitlesName() As Boolean
Grid.ColHidden(Grid.ColIndex("TotalAdvance")) = False
getTitlesName = True
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim SearchFiled As String
    Dim str As String
    Dim ColumnName As String
    Dim i As Integer
    sql = "select * from mofrad order by id  "
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
 
        For i = 1 To rs.RecordCount
            FixedOrChanged(i) = IIf(IsNull(rs("FixedOrChanged").value), 0, rs("FixedOrChanged").value)
            AddOrDiscount(i) = IIf(IsNull(rs("AddOrDiscount").value), 0, rs("AddOrDiscount").value)
            ViewComp(i) = IIf(IsNull(rs("ViewComp").value), False, rs("ViewComp").value)
            Account_code(i) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
             Account_code1(i) = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
             Account_code1(i) = IIf(IsNull(rs("Account_Code1").value), "", rs("Account_Code1").value)
            showMofradAll(i) = IIf(IsNull(rs("showMofradAll").value), False, rs("showMofradAll").value)
            culc30orRminder(i) = IIf(IsNull(rs("culc30orRminder").value), 0, rs("culc30orRminder").value)
      '      If Account_Code(i) = "" Then
      ''      MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & ViewComp(i), vbCritical
       '     getTitlesName = False
       '     Exit Function
       '     End If
            
            
            ZmamAccount(i) = IIf(IsNull(rs("ZmamAccount").value), 0, rs("ZmamAccount").value)
            AdvPaymentdAccount(i) = IIf(IsNull(rs("AdvPaymentdAccount").value), 0, rs("AdvPaymentdAccount").value)
            
            

            
            
              'AdvPaymentdAccount
            If SystemOptions.UserInterface = ArabicInterface Then
                componentname(i) = IIf(IsNull(rs("name").value), "", rs("name").value)
            Else
                componentname(i) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
            End If
             
             
         '   If ViewComp(i) = True And Account_Code(i) = "" And (ZmamAccount(i) <> "True" And AdvPaymentdAccount(i) <> "True") Then
         '   MsgBox " ·„ Ì „ —»ÿ «·Õ”«» «·Œ«’ » " & componentname(i), vbCritical
         '   getTitlesName = False
          
           ' Unload Me
         '     Exit Function
         '   End If
              
              
            With Me.Grid
             
                ColumnName = "Comp" & i

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(0, .ColIndex(ColumnName)) = IIf(IsNull(rs("name").value), "", rs("name").value)
                Else
                    .TextMatrix(0, .ColIndex(ColumnName)) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                End If
                     
                If ViewComp(i) = True Then
                    .ColHidden(.ColIndex(ColumnName)) = False
                Else
                    .ColHidden(.ColIndex(ColumnName)) = True
                End If
                     
            End With
             
 
             
            rs.MoveNext
             
        Next i
  
    End If
 
    rs.Close
End Function

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
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
             Msg = " Select Employee ..!! "
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboEmpName.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
        If val(Total) < 0 Then
     '       Msg = "«·„” Õﬁ«  «’€— „‰ ’›— ·«Ì„ﬂ‰ «·Õ›Ÿ..!! "
     '       MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
     '
     '       SendKeys "{F4}"
     '       Exit Sub
        End If

 With Me.VSFlexGrid1

        For i = 1 To .rows - 1

            If .TextMatrix(i, .ColIndex("MofrdID")) <> "" Then
            If .TextMatrix(i, .ColIndex("DeliverDate")) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ Ì—ÃÏ  ”·Ì„ «·⁄Âœ «Ê·«"
            Else
            MsgBox "Can Not Save Please Drive Assest"
            End If
        Exit Sub
        End If
        End If
        Next i
        End With
        
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("TblVocationEntitlements", "ID", "", True))
         
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
        
            StrSQL = "Delete From TblVocationEntitlementsDet Where VoEntID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblInforVacatiom Where VacatioID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
                  StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords


        
        
        End If
           rs("ID").value = val(XPTxtID.text)
           rs("LastBalanceMonth").value = val(LastBalanceMonth.text)
           
           rs("RecordDate").value = XPDtbTrans.value
           rs("stratDate").value = stratDate.value
           rs("EndDate").value = EndDate.value
           rs("stratDateH").value = ToHijriDate(stratDate.value)
           rs("EndDateH").value = ToHijriDate(EndDate.value)
                  If chkGE.value = Checked Then
            rs("chkGE").value = 1
        Else
            rs("chkGE").value = 0
        End If


           rs("DateSta").value = DateSta.value
           rs("OpretotID").value = IIf(Opretot.BoundText = "", Null, Opretot.BoundText)
           rs("UserID").value = IIf(DCboUserName.BoundText = "", Null, DCboUserName.BoundText)
           rs("EmpID").value = IIf(DcboEmpName.BoundText = "", Null, DcboEmpName.BoundText)
           rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Dcbranch.BoundText)
           rs("JobID").value = IIf(Me.DcboJobsType.BoundText = "", Null, DcboJobsType.BoundText)
           rs("DeptID").value = IIf(Me.DcbDept.BoundText = "", Null, DcbDept.BoundText)
           rs("BignDate").value = BignDate.value
           rs("LastVocatinDate").value = LastVocatinDate.value
           rs("ContDay").value = IIf(Me.TxtContDay.text = "", 0, val(TxtContDay.text))
           rs("LastDayVoc").value = IIf(Me.TxtLastDayVoc.text = "", 0, val(TxtLastDayVoc.text))
           rs("TotalDay").value = IIf(Me.TxtTotalDay.text = "", 0, val(TxtTotalDay.text))
           rs("NoDay").value = IIf(Me.TxtDay.text = "", 0, val(TxtDay.text))
           rs("NoMonth").value = IIf(Me.TxtMonth.text = "", 0, val(TxtMonth.text))
           rs("NoYear").value = IIf(Me.TxtYear.text = "", 0, val(TxtYear.text))
           rs("Remark").value = IIf(Me.TxtRemark.text = "", Null, TxtRemark.text)
           rs("DaySalary").value = IIf(Me.TxtDaySalary.text = "", 0, val(TxtDaySalary.text))
           rs("Salary").value = IIf(Me.TxtSalary.text = "", 0, val(TxtSalary.text))
           rs("DayIncrease").value = IIf(Me.TxtDayIncrease.text = "", 0, val(TxtDayIncrease.text))
           rs("Increase").value = IIf(Me.TxtIncrease.text = "", 0, val(TxtIncrease.text))
           rs("Decrease").value = IIf(Me.TxtDecrease.text = "", 0, val(TxtDecrease.text))
           rs("IDes").value = IDes.text
           rs("TxtNoVaction").value = IIf(Me.TxtNoVaction.text = "", 0, val(TxtNoVaction.text))
           rs("PaymentRecommended").value = IIf(Me.txtPaymentRecommended.text = "", 0, val(txtPaymentRecommended.text))
           rs("DaysCountPay").value = IIf(Me.txtDaysCountPay.text = "", 0, val(txtDaysCountPay.text))
           
           
           '''/////////
           rs("TxtDay2").value = IIf(Me.TxtDay2.text = "", 0, val(TxtDay2.text))
           rs("TxtMonth2").value = IIf(Me.TxtMonth2.text = "", 0, val(TxtMonth2.text))
           rs("TxtYear2").value = IIf(Me.TxtYear2.text = "", 0, val(TxtYear2.text))
           
           rs("TxtDay3").value = IIf(Me.TxtDay3.text = "", 0, val(TxtDay3.text))
           rs("TxtMonth3").value = IIf(Me.TxtMonth3.text = "", 0, val(TxtMonth3.text))
           rs("TxtYear3").value = IIf(Me.TxtYear3.text = "", 0, val(TxtYear3.text))
           rs("TxtAddDay").value = IIf(Me.TxtAddDay.text = "", 0, val(TxtAddDay.text))
           rs("TxtDiscouDay").value = IIf(Me.TxtDiscouDay.text = "", 0, val(TxtDiscouDay.text))
           ''//////////
           rs("InsuranceValue").value = IIf(Me.TxtInsuranceValue.text = "", 0, val(TxtInsuranceValue.text))
           rs("GetInsurance").value = IIf(Me.TxtGetInsurance.text = "", 0, val(TxtGetInsurance.text))
         '  rs("NoMonth").value = IIf(Me.TxtNoMonth.text = "", 0, val(TxtNoMonth.text))
           
           rs("DaySalVocation").value = IIf(Me.TxtDaySalVocation.text = "", 0, val(TxtDaySalVocation.text))
           rs("SalaryVocation").value = IIf(Me.TxtSalaryVocation.text = "", 0, val(TxtSalaryVocation.text))
           rs("DayEntitOther").value = IIf(Me.TxtDayEntitOther.text = "", 0, val(TxtDayEntitOther.text))
           rs("SalEntitOther").value = IIf(Me.TxtSalEntitOther.text = "", 0, val(TxtSalEntitOther.text))
           
           rs("ADDACC").value = ADDACC.BoundText
           rs("DISACC").value = DISACC.BoundText
           
           
           
           rs("Other").value = IIf(Me.TxtOther.text = "", 0, val(TxtOther.text))
           rs("Advance").value = IIf(Me.TxtAdvance.text = "", 0, val(TxtAdvance.text))
           rs("ValueTickt").value = IIf(Me.TxtValueTickt.text = "", 0, val(TxtValueTickt.text))
           ''//////02 09 2015
           rs("YaerOut").value = IIf(Me.TxtYaerOut.text = "", 0, val(TxtYaerOut.text))
           rs("MontOut").value = IIf(Me.TxtMontOut.text = "", 0, val(TxtMontOut.text))
           rs("DayOut").value = IIf(Me.TxtVSa.text = "", 0, val(TxtVSa.text))
           rs("DayAbs").value = IIf(Me.TxtDayAbs.text = "", 0, val(TxtDayAbs.text))
           rs("MoAbs").value = IIf(Me.TxtMoAbs.text = "", 0, val(TxtMoAbs.text))
           rs("YearAbs").value = IIf(Me.TxtYearAbs.text = "", 0, val(TxtYearAbs.text))
           rs("ToalAbsent").value = IIf(Me.TxtToalAbsent.text = "", 0, val(TxtToalAbsent.text))
           rs("DuVocation").value = IIf(Me.TxtDuVocation.text = "", 0, val(TxtDuVocation.text))
           rs("LastTotal").value = IIf(Me.Total.text = "", 0, val(Total.text))
            rs("TypEndService").value = IIf(val(Me.dctype.BoundText) = 0, 0, val(dctype.BoundText))
            
            rs("WithoutSala1").value = IIf(Me.TxtWithOutSala1.text = "", 0, val(TxtWithOutSala1.text))
           rs("NoOrder").value = IIf(Me.Txtorder.text = "", 0, val(Txtorder.text))
           rs("NoVacation").value = IIf(Me.TxtNoVacation.text = "", 0, val(TxtNoVacation.text))
           rs("BasedOn").value = IIf(Me.CbBasedOn.ListIndex = -1, Null, val(CbBasedOn.ListIndex))
           rs("NewAbsent").value = IIf(Me.TxtNewAbsent.text = "", 0, val(TxtNewAbsent.text))
           '''/////////
          If Opt(1).value = True Then
          rs("Vact_Work").value = 1
          ElseIf Opt(2).value = True Then
          rs("Vact_Work").value = 2
          
          End If
         
          '  rs("TotalDue").value = IIf(Me.TxtTolaMostak.text = "", 0, val(TxtTolaMostak.text))
          '  rs("NetDue").value = IIf(Me.NetTotal.text = "", 0, val(NetTotal.text))
          '  rs("TotalCut").value = IIf(Me.TxtTotalCut.text = "", 0, val(TxtTotalCut.text))
          '  rs("NetTotal").value = IIf(Me.Total.text = "", 0, val(Total.text))
           ''////////////////////
           If Ch(0).value = vbChecked Then
           rs("ch0").value = 1
          Else
           rs("ch0").value = 0
         End If
         If Ch(1).value = vbChecked Then
           rs("ch1").value = 1
         Else
           rs("ch1").value = 0
         End If
         If Ch(2).value = vbChecked Then
           rs("ch2").value = 1
         Else
           rs("ch2").value = 0
         End If
         If Ch(3).value = vbChecked Then
           rs("ch3").value = 1
         Else
           rs("ch3").value = 0
         End If
          If Ch(4).value = vbChecked Then
           rs("ch4").value = 1
         Else
           rs("ch4").value = 0
         End If
        If Ch(5).value = vbChecked Then
           rs("ch5").value = 1
         Else
           rs("ch5").value = 0
         End If
       If Ch(6).value = vbChecked Then
           rs("ch6").value = 1
         Else
           rs("ch6").value = 0
        End If
       If Ch(7).value = vbChecked Then
           rs("ch7").value = 1
         Else
           rs("ch7").value = 0
         End If
        If Ch(8).value = vbChecked Then
           rs("ch8").value = 1
         Else
           rs("ch8").value = 0
         End If
         
            If Ch(9).value = vbChecked Then
           rs("ch9").value = 1
          Else
           rs("ch9").value = 0
         End If
         
         rs("PreSalary").value = val(TxtPreSalary.text)
           '//////////
         If ChBooked.value = vbChecked Then
           rs("Booked").value = 1
         Else
           rs("Booked").value = 0
         End If
          If ChDelivery.value = vbChecked Then
           rs("Delivery").value = 1
         Else
           rs("Delivery").value = 0
         End If
         ''' aladein ADD
         If Option1.value = True Then
         rs("Chekk").value = 0
         Else
         rs("Chekk").value = 1
         End If
         rs.update
         If val(TxtWithOutSala1.text) <> 0 Then
         SaveInformationVacation 0, val(DcboEmpName.BoundText), val(TxtWithOutSala1.text)
         End If
         If val(TxtNewAbsent.text) <> 0 Then
         SaveInformationVacation 1, val(DcboEmpName.BoundText), val(TxtNewAbsent.text)
         End If
         ''///
          Cn.Execute "Update TblVocation set FlagPayed=1  where ID =" & val(Txtorder.text) & " "
          If val(TxtNoVaction.text) > 0 Then
          Cn.Execute "Update TblEmbarkation set VacationPaied=1  where TypeVacation=1 and Emp_ID=" & val(Me.DcboEmpName.BoundText) & " and  ID in(" & Me.IDes.text & ") "
         End If
                  Set RsDev = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.TblVocationEntitlementsDet Where (1 = -1)"
         RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
      With Me.Fg

        For i = 1 To .rows - 1

            If .TextMatrix(i, .ColIndex("MofrdID")) <> "" Then
         
                RsDev.AddNew
                RsDev("VoEntID").value = val(Me.XPTxtID.text)
                RsDev("MofrdID").value = val(.TextMatrix(i, .ColIndex("MofrdID")))
              
                RsDev("Valu").value = val(.TextMatrix(i, .ColIndex("Valu")))
     
                RsDev("TypeM").value = 0
                RsDev.update
                    
            End If
            
            '
        Next i

    End With

    RsDev.Close
    ''//
           
         Set RsDev = New ADODB.Recordset
         StrSQL = "SELECT     *  from dbo.TblVocationEntitlementsDet Where (1 = -1)"
         RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      With Me.VSFlexGrid1

        For i = 1 To .rows - 1

            If .TextMatrix(i, .ColIndex("MofrdID")) <> "" Then
         
                RsDev.AddNew
                RsDev("VoEntID").value = val(Me.XPTxtID.text)
                RsDev("MofrdID").value = val(.TextMatrix(i, .ColIndex("MofrdID")))
                
                RsDev("EmpID").value = val(.TextMatrix(i, .ColIndex("EmpID")))
                 RsDev("DeliverDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("DeliverDate"))), .TextMatrix(i, .ColIndex("DeliverDate")), Null)
                RsDev("ReciveDate").value = IIf(IsDate(.TextMatrix(i, .ColIndex("ReciveDate"))), .TextMatrix(i, .ColIndex("ReciveDate")), Null)
                RsDev("TypeM").value = 1
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
 If Ch(8).value = vbChecked Then
    With Grid1
        For i = .FixedRows To .rows - 1
            If .TextMatrix(i, .ColIndex("RecordDate")) <> "" Then
                If .cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                    Cn.Execute " Update emp_salary set Payed=1, VocEntitID=" & val(Me.XPTxtID.text) & " where RecordDate=" & SQLDate(.TextMatrix(i, .ColIndex("RecordDate")), True) & " and emp_id=" & val(Me.DcboEmpName.BoundText) & ""
                Else
                    Cn.Execute " Update emp_salary set Payed=0, VocEntitID=" & val(Me.XPTxtID.text) & " where RecordDate=" & SQLDate(.TextMatrix(i, .ColIndex("RecordDate")), True) & " and emp_id=" & val(Me.DcboEmpName.BoundText) & ""
                    'If Change_filed_value(val(.TextMatrix(i, .ColIndex("id"))), "id", "Payed", "emp_salary", 1) Then
                    'End If
                End If
            End If
        Next i
    End With
  End If
  SaveSalary
    Ele(10).Visible = True
C1Elastic2.Visible = False
Ele(12).Visible = False
        Cn.CommitTrans
        BeginTrans = False
       'sa RsDev.Close
        Set RsDev = Nothing
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
              If SystemOptions.UserInterface = ArabicInterface Then
              Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                Msg = "This Record Saved "
                Msg = Msg & " You Need To Enter Another Record  "
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
               If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
               Else
               MsgBox " Saved Successfully  ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
               End If
        End Select
        Retrive
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

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
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
Function ChePayment() As Boolean
Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
sql = "Select * from Notes where Due=" & val(XPTxtID.text) & " and NoteType=5 and CashingType=8 "
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
ChePayment = True
Else
ChePayment = False
End If
End Function



Function CheWork() As Boolean
Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
sql = "Select * from TblEmbarkation where Emp_ID=" & val(DcboEmpName.BoundText) & " and stratDate=" & SQLDate(stratDate.value, True) & " and EndDate=" & SQLDate(EndDate.value, True) & ""
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
CheWork = True
Else
CheWork = False
End If
End Function
Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    Dim i As Integer
'    On Error GoTo ErrTrap
    If CheWork() = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… „— »ÿ… »„»«‘—… ⁄„·"
            Else
            MsgBox "Can not delete this process Linked to the initiation of work"
            End If
            Exit Sub
            End If
          If ChePayment() = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… „— »ÿ… »«·„œ›Ê⁄« "
            Else
            MsgBox "Can not delete this process Linked to Payments"
            End If
            Exit Sub
            End If
          
    If XPTxtID.text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
Msg = "Confirm Delete"
End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
        
                  If Opt(0).value = True Then
           StrSQL = "Update TblEmployee Set  jopstatusid=1 ,workstate=1 Where Emp_ID=" & val(DcboEmpName.BoundText) & ""
              Cn.Execute StrSQL, , adExecuteNoRecords
         End If
                rs.delete
                  Deletepost Me.Name, "TblVocationEntitlements", "Id", val(Me.DcbDept.BoundText), val(Dcbranch.BoundText), val(XPTxtID.text), XPTxtID
    
    With Grid1
        For i = .FixedRows To .rows - 1
            If .TextMatrix(i, .ColIndex("RecordDate")) <> "" Then
            Cn.Execute " Update emp_salary set Payed=Null, VocEntitID=Null where RecordDate=" & SQLDate(.TextMatrix(i, .ColIndex("RecordDate")), True) & " and emp_id=" & val(Me.DcboEmpName.BoundText) & ""
            End If
        Next i
    End With
       If val(TxtNoVaction.text) > 0 Then
          Cn.Execute "Update TblEmbarkation set VacationPaied=null  where TypeVacation=1 and Emp_ID=" & val(Me.DcboEmpName.BoundText) & " and  ID in(" & Me.IDes.text & ") "
         End If
         
           Cn.Execute "Delete from TblVacationSalary where VacationID=" & val(XPTxtID.text) & ""
            Cn.Execute "Update TblVocation set FlagPayed=null  where ID =" & val(Txtorder.text) & " "
                StrSQL = "Delete From TblVocationEntitlements Where ID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblVocationEntitlementsDet Where VoEntID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                     StrSQL = "Delete From TblInforVacatiom Where VacatioID=" & val(Me.XPTxtID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords


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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This Process is not Availablet Because  there was no Records"
        End If
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
'
' Dim sql As String
'  Dim Rs1 As New ADODB.Recordset
' Dim i As Integer
'    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
'  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
''  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
 ' sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
 ' sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
'sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.name & "')"
'sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "
'
'    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Rs1.RecordCount > 0 Then
'            currentdate = Now
'            For i = 1 To Rs1.RecordCount
'              RSApproval.AddNew
''                RSApproval("ScreenName").value = Me.name
 '               RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
 '              RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
 ''               RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
  '               RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
  '                RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
  '                 RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
  ''              RSApproval("Transaction_Date").value = Date
   '
   '               RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.name), currentdate)
   '            RSApproval("SendTime").value = currentdate
'
'                 If i = 1 Then
'                        RSApproval("Currcursor").value = 1
'                         RSApproval("FromUser").value = user_name
'                End If
'
'                RSApproval.update
'                Rs1.MoveNext
'            Next i
'
'    End If
'
'
'
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
'
'    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
' If Not (RsDetails.EOF Or RsDetails.BOF) Then
'        GRID2.Rows = RsDetails.RecordCount + 1
'
'
'        For Num = 1 To RsDetails.RecordCount
'
'       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
'    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
'   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
'   Else
'    GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
'    End If
'
'        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
'           If SystemOptions.UserInterface = ArabicInterface Then
'            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
'          Else
'             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
'          End If
'            If SystemOptions.UserInterface = ArabicInterface Then
'            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
'            Else
'            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
'            End If
'            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
'          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
'
'
'RsDetails.MoveNext
'If Num = RsDetails.RecordCount Then
'
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
'
'End If
'
'        Next Num
'Else
' GRID2.Rows = 1
'    End If
''RsDetails.Close
'
'End Function
Sub RtriverAsse(Optional EmpID As Integer = 0)
Dim sql As String
Dim i As Integer
Dim RsDev As ADODB.Recordset
sql = " SELECT     TOP 100 PERCENT dbo.TblAssestes.AsID, dbo.TblAssestes.AsName, dbo.TblAssestes.AsCode, TblEmployee_2.Emp_Name, TblEmployee_2.Fullcode,"
sql = sql & "                      TblEmployee_2.Emp_Namee, dbo.TblEmpAsest.ToEmId, TblEmployee_1.Emp_Name AS Emp_NameTo, TblEmployee_1.Fullcode AS FullcodeTo,"
sql = sql & "                       TblEmployee_1.Emp_Namee AS Emp_NameToE, dbo.TblEmpAsest.DeliverDate, dbo.TblEmpAsest.PostedDate, dbo.TblEmpAsestDetails.Qunt,"
sql = sql & "                       dbo.TblEmpAsestDetails.DIFF , dbo.TblEmpAsestDetails.FlagAs, dbo.TblEmpAsest.TypeAsset, dbo.TblEmpAsest.EmpAsestID"
sql = sql & "  FROM         dbo.TblEmpAsest LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmpAsestDetails ON dbo.TblEmpAsest.EmpAsID = dbo.TblEmpAsestDetails.IDAseset LEFT OUTER JOIN"
sql = sql & "                       dbo.TblAssestes ON dbo.TblEmpAsestDetails.AsID = dbo.TblAssestes.AsID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmployee TblEmployee_1 ON dbo.TblEmpAsest.ToEmId = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblEmployee TblEmployee_2 ON dbo.TblEmpAsest.EmpAsestID = TblEmployee_2.Emp_ID"
sql = sql & "  Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblEmpAsest.EmpAsestID =" & EmpID & ")"
Set RsDev = New ADODB.Recordset
       RsDev.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
 VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
           VSFlexGrid1.rows = 1
    If (RsDev.RecordCount > 0) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid1
    
            .rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .rows - 1
 
                .TextMatrix(i, .ColIndex("MofrdID")) = IIf(IsNull(RsDev("AsID").value), "", RsDev("AsID").value)
            
                .TextMatrix(i, .ColIndex("AsCode")) = IIf(IsNull(RsDev("AsCode").value), "", RsDev("AsCode").value)
                .TextMatrix(i, .ColIndex("DeliverDate")) = IIf(IsNull(RsDev("DeliverDate").value), "", RsDev("DeliverDate").value)
                .TextMatrix(i, .ColIndex("ReciveDate")) = IIf(IsNull(RsDev("PostedDate").value), "", RsDev("PostedDate").value)
            
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(RsDev("ToEmId").value), "", RsDev("ToEmId").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("AsName").value), "", RsDev("AsName").value)
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_NameTo").value), "", RsDev("Emp_NameTo").value)
                Else
                 .TextMatrix(i, .ColIndex("mofrd")) = IIf(IsNull(RsDev("AsName").value), "", RsDev("AsName").value)
                .TextMatrix(i, .ColIndex("Emp_NameTo")) = IIf(IsNull(RsDev("Emp_NameToE").value), "", RsDev("Emp_NameToE").value)
                End If
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
End Sub
Sub GetProInsurance()
Dim sql As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
sql = "select * from TblSocialInsurance order by  ID "
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then

    TxtStay.text = IIf(IsNull(Rs6("ResidentVal1").value), 0, Rs6("ResidentVal1").value)
    TxtCivilin.text = IIf(IsNull(Rs6("CitizenVal1").value), 0, Rs6("CitizenVal1").value)

Else

    TxtStay.text = 0
    TxtCivilin.text = 0
End If
End Sub

Function GetInsurnceValue() As Double
Dim InsTotal As Double
If Me.TxtModFlg.text <> "R" Then
Dim sql As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
 sql = "  SELECT     TOP 100 PERCENT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
 sql = sql & "                     SUM(dbo.EmpSalaryComponent.[Value]) AS Salary, dbo.mofrad.Insurances, dbo.jopstatus.Insurances AS InsurancesJob, dbo.jopstatus.resignationInt,"
 sql = sql & "                     dbo.TblEmployee.InsuranceState, dbo.TblEmployee.NationalityE, dbo.TblEmployee.Nationality, dbo.TblEmployee.InstanceDateM, dbo.TblEmployee.BranchId,"
 sql = sql & "                      dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE"
 sql = sql & " FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
 sql = sql & "                     dbo.TblEmployee ON dbo.TblBranchesData.branch_id = dbo.TblEmployee.BranchId LEFT OUTER JOIN"
 sql = sql & "                     dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id LEFT OUTER JOIN"
 sql = sql & "                     dbo.mofrad RIGHT OUTER JOIN"
 sql = sql & "                     dbo.EmpSalaryComponent ON dbo.mofrad.id = dbo.EmpSalaryComponent.mofrad_type ON dbo.TblEmployee.Emp_ID = dbo.EmpSalaryComponent.emp_ID"
 sql = sql & "  WHERE     ((DATEPART(year, dbo.EmpSalaryComponent.EntIncresDataM) < " & year(stratDate.value) & ") OR"
 sql = sql & "                     (DATEPART(year, dbo.EmpSalaryComponent.EntIncresDataM) IS NULL))"
sql = sql & " and dbo.TblEmployee.Emp_ID=" & val(Me.DcboEmpName.BoundText) & ""
  
 sql = sql & " GROUP BY dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.mofrad.Insurances,"
 sql = sql & "                     dbo.jopstatus.Insurances, dbo.jopstatus.resignationInt, dbo.TblEmployee.InsuranceState, dbo.TblEmployee.NationalityE, dbo.TblEmployee.Nationality,"
 sql = sql & "                     dbo.TblEmployee.InstanceDateM , dbo.TblEmployee.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
 sql = sql & " HAVING      (dbo.mofrad.Insurances = 1) AND (dbo.jopstatus.Insurances = 1) AND (dbo.jopstatus.resignationInt IS NULL) AND (dbo.TblEmployee.InsuranceState = 1) AND"
 sql = sql & "                     (dbo.TblEmployee.InstanceDateM <= " & SQLDate(Me.stratDate.value, True) & " OR"
 sql = sql & "                     dbo.TblEmployee.InstanceDateM IS NULL) OR"
 sql = sql & "                     (dbo.jopstatus.resignationInt <> 2) AND (dbo.jopstatus.resignationInt <> 1)"
 sql = sql & " ORDER BY dbo.TblEmployee.Emp_ID"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
 If Rs5("Nationality").value = "”⁄ÊœÌ" Or Rs5("Nationality").value = "”⁄ÊœÏ" Or Rs5("Nationality").value = "Saudi" Then
InsTotal = ((IIf(IsNull(Rs5("salary").value), 0, Rs5("salary").value) * val(TxtCivilin.text)) / 100)
Else
InsTotal = ((IIf(IsNull(Rs5("salary").value), 0, Rs5("salary").value) * val(TxtStay.text)) / 100)
End If
InsTotal = Round(InsTotal, 2)
GetInsurnceValue = InsTotal
Else
GetInsurnceValue = 0
End If
End If
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

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "    „” Õﬁ«  «·ﬁÌ«„ »«Ã«“…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   „” Õﬁ«  «·ﬁÌ«„ »«Ã«“…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   „” Õﬁ«  «·ﬁÌ«„ »«Ã«“…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   „” Õﬁ«  «·ﬁÌ«„ »«Ã«“…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   „” Õﬁ«  «·ﬁÌ«„ »«Ã«“…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   „” Õﬁ«  «·ﬁÌ«„ »«Ã«“…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "   „” Õﬁ«  «·ﬁÌ«„ »«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   „” Õﬁ«  «·ﬁÌ«„ »«Ã«“…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   „” Õﬁ«  «·ﬁÌ«„ »«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   „” Õﬁ«  «·ﬁÌ«„ »«Ã«“… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "   „” Õﬁ«  «·ﬁÌ«„ »«Ã«“…  ", 1, 15204351, -2147483630
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

 

Private Sub XPDtbTrans_Click()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.text = ""
End If

End Sub

