VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmNewGard10 
   Caption         =   "Ã—œ «·«’Ê· «·À«» …"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13485
   HelpContextID   =   90
   Icon            =   "FrmNewGard10.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmNewGard10.frx":038A
   RightToLeft     =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   13485
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   7455
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   13485
      _cx             =   23786
      _cy             =   13150
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
      AutoSizeChildren=   8
      BorderWidth     =   1
      ChildSpacing    =   1
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
      GridRows        =   5
      GridCols        =   6
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmNewGard10.frx":0714
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   450
         Index           =   5
         Left            =   15
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   6990
         Width           =   13455
         _cx             =   23733
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
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.TextBox XPTxtSum 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   15045
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   390
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   3600
            TabIndex        =   12
            Top             =   75
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ «·›⁄·ÌÂ"
            Height          =   330
            Index           =   63
            Left            =   8160
            TabIndex        =   65
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label LblTotalQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Left            =   7230
            TabIndex        =   64
            Top             =   0
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label lblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   14820
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   60
            Width           =   1695
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·—’Ìœ"
            Height          =   270
            Index           =   3
            Left            =   16155
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   120
            Width           =   1650
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   330
            Index           =   6
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   105
            Width           =   870
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   135
            Width           =   825
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   285
            Left            =   2265
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   105
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   375
            Index           =   5
            Left            =   1215
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   0
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   495
            Index           =   4
            Left            =   2985
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   0
            Width           =   585
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4935
         Index           =   3
         Left            =   15
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2040
         Width           =   13455
         _cx             =   23733
         _cy             =   8705
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
         AutoSizeChildren=   8
         BorderWidth     =   2
         ChildSpacing    =   1
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
         GridRows        =   3
         GridCols        =   3
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmNewGard10.frx":07A7
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin MSComctlLib.Toolbar TBr 
            Height          =   630
            Left            =   495
            TabIndex        =   26
            Top             =   4545
            Width           =   12465
            _ExtentX        =   21987
            _ExtentY        =   1111
            ButtonWidth     =   609
            ButtonHeight    =   1005
            Appearance      =   1
            _Version        =   393216
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   690
            Index           =   4
            Left            =   12975
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   450
            _cx             =   794
            _cy             =   1217
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
            Begin VB.TextBox TxtPrice 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   30
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   300
               Width           =   45
            End
            Begin VB.TextBox TxtSerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   135
               MaxLength       =   20
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   300
               Width           =   75
            End
            Begin VB.TextBox TxtQuantity 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   75
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   300
               Width           =   60
            End
            Begin VB.ComboBox CboItemCase 
               Height          =   315
               Left            =   210
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   300
               Width           =   60
            End
            Begin MSDataListLib.DataCombo DCboItemsName 
               Height          =   315
               Left            =   270
               TabIndex        =   1
               Top             =   270
               Width           =   90
               _ExtentX        =   159
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboItemsCode 
               Height          =   315
               Left            =   360
               TabIndex        =   0
               Top             =   300
               Width           =   90
               _ExtentX        =   159
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdAdd 
               Height          =   420
               Left            =   0
               TabIndex        =   6
               Top             =   210
               Width           =   30
               _ExtentX        =   53
               _ExtentY        =   741
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
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
               ButtonImage     =   "FrmNewGard10.frx":0809
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               LowerToggledContent=   0   'False
               ColorTextShadow =   -2147483637
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”⁄—"
               Height          =   270
               Index           =   26
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   30
               Width           =   45
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   270
               Index           =   27
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   30
               Width           =   45
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ì—Ì«·"
               Height          =   390
               Index           =   28
               Left            =   135
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   30
               Width           =   75
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·’‰›"
               Height          =   270
               Index           =   29
               Left            =   210
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   30
               Width           =   60
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈”„ «·’‰›"
               Height          =   270
               Index           =   30
               Left            =   285
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   30
               Width           =   75
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   270
               Index           =   31
               Left            =   375
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   30
               Width           =   75
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   360
            Left            =   30
            TabIndex        =   63
            Top             =   4545
            Visible         =   0   'False
            Width           =   13395
            _cx             =   23627
            _cy             =   635
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
            Rows            =   2
            Cols            =   22
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmNewGard10.frx":0BA3
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
            WallPaperAlignment=   0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
            Height          =   4875
            Left            =   30
            TabIndex        =   102
            Top             =   30
            Width           =   13395
            _cx             =   23627
            _cy             =   8599
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
            Rows            =   2
            Cols            =   26
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmNewGard10.frx":0F34
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
            WallPaperAlignment=   0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label LblItemsCount 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   4545
            Width           =   450
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1380
         Index           =   1
         Left            =   15
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   645
         Width           =   13455
         _cx             =   23733
         _cy             =   2434
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
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬂ· «·«’Ê·"
            Height          =   375
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CommandButton Command3 
            Caption         =   "”‰œ«   «·“Ì«œÂ"
            Height          =   315
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   1080
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Frame Frame2 
            Caption         =   "ŒÌ«—«  ÿ»«⁄Â «·Ã—œ"
            Height          =   1335
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   0
            Width           =   3255
            Begin VB.CommandButton CMdPrit 
               Caption         =   "ÿ»«⁄Â"
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   840
               Width           =   1695
            End
            Begin VB.Frame Frame3 
               Caption         =   "⁄—÷ «·”⁄—"
               Height          =   735
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   120
               Visible         =   0   'False
               Width           =   1815
               Begin VB.OptionButton Opt3 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”⁄—«·»Ì⁄"
                  Height          =   195
                  Index           =   1
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.OptionButton Opt3 
                  Alignment       =   1  'Right Justify
                  Caption         =   "”⁄—«· ﬂ·›…"
                  Height          =   195
                  Index           =   0
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   240
                  Width           =   1095
               End
            End
            Begin VB.OptionButton opt2 
               Alignment       =   1  'Right Justify
               Caption         =   "«·“Ì«œ… ›ﬁÿ"
               Height          =   195
               Index           =   2
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   720
               Width           =   1335
            End
            Begin VB.OptionButton opt2 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄Ã“ ›ﬁÿ"
               Height          =   195
               Index           =   1
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   480
               Width           =   1335
            End
            Begin VB.OptionButton opt2 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬂ· «·Ã—œ"
               Height          =   195
               Index           =   0
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   9960
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "”‰œ«   «·⁄Ã“"
            Height          =   315
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   1080
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   15000
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   1080
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   16560
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton CMDStartSetelment 
            Caption         =   " ‰›Ì– «· ”ÊÌ«  «·Ã—œÌÂ"
            Height          =   315
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   720
            Width           =   2175
         End
         Begin VB.CommandButton CMDStartGard 
            Caption         =   " ‰›Ì– «·Ã—œ"
            Height          =   315
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   360
            Width           =   2175
         End
         Begin VB.Frame Frame1 
            Caption         =   "Õœœ ÿ—Ìﬁ… «·«œŒ«·"
            Height          =   1095
            Left            =   15750
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   480
            Visible         =   0   'False
            Width           =   3315
            Begin VB.CommandButton Command1 
               Caption         =   " ÕœÌœ «·„·›..."
               Height          =   255
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton OPT 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÌœÊÌ"
               Height          =   195
               Index           =   2
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   720
               Width           =   1575
            End
            Begin VB.OptionButton OPT 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂ· «’‰«› «·„Œ“‰"
               Height          =   195
               Index           =   1
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   480
               Width           =   1575
            End
            Begin VB.OptionButton OPT 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ „·›"
               Height          =   195
               Index           =   0
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.TextBox txtopening_balance_voucher_id 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2925
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1560
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1365
            Left            =   -7920
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   0
            Width           =   8010
            Begin MSDataListLib.DataCombo DcboDebitSide 
               Height          =   315
               Left            =   1680
               TabIndex        =   37
               Top             =   180
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCreditSide 
               Height          =   315
               Left            =   1680
               TabIndex        =   38
               Top             =   510
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label LblAccountInterval 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   5580
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   510
               Width           =   1095
            End
            Begin VB.Label LblDevID 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   5580
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·› —… :"
               Height          =   285
               Index           =   9
               Left            =   6690
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   510
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ﬁÌœ:"
               Height          =   285
               Index           =   8
               Left            =   6690
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   180
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› œ«∆‰"
               Height          =   285
               Index           =   7
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   510
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› „œÌ‰"
               Height          =   285
               Index           =   32
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   180
               Width           =   885
            End
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   105
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   1170
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   345
            Left            =   11250
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   60
            Width           =   1350
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   1020
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   630
            Visible         =   0   'False
            Width           =   825
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   345
            Left            =   8400
            TabIndex        =   30
            Top             =   0
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   93388803
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   9960
            TabIndex        =   61
            Top             =   840
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   13650
            TabIndex        =   66
            Top             =   510
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   1065
            Index           =   2
            Left            =   7410
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   360
            Width           =   2355
            _cx             =   4154
            _cy             =   1879
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   " «·› —… «·“„‰Ì…"
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
            Style           =   1
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            Frame           =   0
            FrameStyle      =   5
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin MSComCtl2.DTPicker DTPickerAccFrom 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11265
                  SubFormatType   =   3
               EndProperty
               Height          =   345
               Left            =   90
               TabIndex        =   69
               ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
               Top             =   240
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   -2147483624
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   93388803
               CurrentDate     =   37357
            End
            Begin MSComCtl2.DTPicker DTPickerAccTo 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11265
                  SubFormatType   =   3
               EndProperty
               Height          =   345
               Left            =   90
               TabIndex        =   70
               ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
               Top             =   600
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   -2147483624
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   93388803
               CurrentDate     =   37357
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   285
               Index           =   11
               Left            =   1590
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   600
               Width           =   555
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   285
               Index           =   10
               Left            =   1590
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   285
               Width           =   555
            End
         End
         Begin MSDataListLib.DataCombo DCAccount1 
            Height          =   315
            Left            =   3840
            TabIndex        =   85
            Top             =   30
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCAccount2 
            Height          =   315
            Left            =   3000
            TabIndex        =   87
            Top             =   870
            Visible         =   0   'False
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   9960
            TabIndex        =   100
            Top             =   480
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «· ”ÊÌÂ »«·“Ì«œ…"
            Height          =   375
            Index           =   13
            Left            =   5715
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   840
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «· ”ÊÌÂ "
            Height          =   375
            Index           =   12
            Left            =   6195
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   0
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘—Ê⁄"
            Height          =   375
            Index           =   2
            Left            =   12405
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   480
            Width           =   945
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   12600
            TabIndex        =   62
            Top             =   840
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   375
            Index           =   0
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   105
            Width           =   600
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”·”·"
            Height          =   375
            Index           =   1
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   75
            Width           =   600
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   615
         Left            =   15
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   15
         Width           =   13455
         _cx             =   23733
         _cy             =   1085
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
         Caption         =   "Ã—œ «·«’Ê· «·À«» …  "
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.CheckBox chkDifferentAccounts 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "«· ”ÊÌÂ ⁄·Ï Õ”«»«  „Œ ·›…"
            Height          =   195
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   360
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox ChkStartGard 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   " „  ‰›Ì– «·Ã—œ"
            Enabled         =   0   'False
            Height          =   195
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkStartSetelment 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   " „  ‰›Ì– «· ”ÊÌ«  «·Ã—œÌÂ"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1755
            TabIndex        =   46
            Top             =   120
            Width           =   885
            _ExtentX        =   1561
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
            ButtonImage     =   "FrmNewGard10.frx":1341
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
            Left            =   975
            TabIndex        =   47
            Top             =   120
            Width           =   780
            _ExtentX        =   1376
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
            ButtonImage     =   "FrmNewGard10.frx":16DB
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
            Left            =   2700
            TabIndex        =   48
            Top             =   120
            Width           =   600
            _ExtentX        =   1058
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
            ButtonImage     =   "FrmNewGard10.frx":1A75
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
            Left            =   150
            TabIndex        =   49
            Top             =   120
            Width           =   795
            _ExtentX        =   1402
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
            ButtonImage     =   "FrmNewGard10.frx":1E0F
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   870
         Index           =   0
         Left            =   15
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   6570
         Visible         =   0   'False
         Width           =   13455
         _cx             =   23733
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   2
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
            Height          =   585
            Index           =   0
            Left            =   12060
            TabIndex        =   51
            Top             =   165
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   1032
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   585
            Index           =   1
            Left            =   10590
            TabIndex        =   52
            Top             =   165
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   1032
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   585
            Index           =   2
            Left            =   9060
            TabIndex        =   53
            Top             =   195
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   1032
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   585
            Index           =   3
            Left            =   7680
            TabIndex        =   54
            Top             =   165
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1032
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   585
            Index           =   4
            Left            =   5790
            TabIndex        =   55
            Top             =   165
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   1032
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   585
            Index           =   5
            Left            =   4545
            TabIndex        =   56
            Top             =   165
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   1032
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   585
            Index           =   6
            Left            =   30
            TabIndex        =   57
            Top             =   165
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1032
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   585
            Index           =   7
            Left            =   2925
            TabIndex        =   58
            Top             =   165
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   1032
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   585
            Left            =   1620
            TabIndex        =   59
            Top             =   165
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   1032
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
      End
   End
End
Attribute VB_Name = "FrmNewGard10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim BalanceReport As ClsOpeningBalanceReport
Dim cSearchDcbo As clsDCboSearch
Dim NewGrid As New ClsGrid
Dim ReportType  As Integer
Dim priceType As Integer

Private Sub chkDifferentAccounts_Click()

    If chkDifferentAccounts.value = vbChecked Then
        FG.Editable = flexEDKbdMouse
    Else
        FG.Editable = flexEDNone

    End If

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CMdPrit_Click()
    GardReport val(Me.XPTxtBillID.Text), ReportType, priceType
End Sub

Private Sub CMDStartGard_Click()
    'If ChkStartGard.value = vbChecked Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '    MsgBox " „ ⁄„· «·Ã—œ „‰ ﬁ»· Ê·« Ì„ﬂ‰ «· ⁄œÌ·"
    '    Else
    '    MsgBox "Cant Do"
    '    End If
    '    Exit Sub
    'End If

    Dim ItemID As Long
    Dim UnitID As Long
    Dim itemsize As Long
    Dim ColorID As Long
    Dim ClassId As Long
Dim ParrtNoCode As String
    With FG

        For RowNum = 1 To FG.Rows - 1
        
            ItemID = val(.TextMatrix(RowNum, FG.ColIndex("Code")))
            UnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            itemsize = val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")))
            ColorID = val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID")))
            ClassId = val(FG.TextMatrix(RowNum, FG.ColIndex("ClassID")))
            ParrtNoCode = (FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")))
       If ParrtNoCode = "" Then
            FG.TextMatrix(RowNum, FG.ColIndex("GardQty")) = GetActualItemQty(val(Me.DCboStoreName.BoundText), Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value, ItemID, UnitID, itemsize, ColorID, ClassId)
       Else
       FG.TextMatrix(RowNum, FG.ColIndex("GardQty")) = GetQtyByBarcode(val(Me.DCboStoreName.BoundText), Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value, ParrtNoCode)
       
       End If
       
            FG.TextMatrix(RowNum, FG.ColIndex("Gardresult")) = val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) - val(FG.TextMatrix(RowNum, FG.ColIndex("GardQty")))
 
            If val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult"))) < 0 Then
                FG.TextMatrix(RowNum, FG.ColIndex("Gardresult2")) = Abs(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult")))
                FG.TextMatrix(RowNum, FG.ColIndex("Gardresult1")) = 0
            ElseIf val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult"))) >= 0 Then
                FG.TextMatrix(RowNum, FG.ColIndex("Gardresult1")) = Abs(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult")))
                FG.TextMatrix(RowNum, FG.ColIndex("Gardresult2")) = 0
            End If
 
        Next RowNum

        FG.AutoSize 0, FG.Cols - 1, False
    End With

    ChkStartGard.value = vbChecked
 
    Cmd_Click (1)
    Cmd_Click (2)
End Sub

Private Sub MinusVoucher()
    ' On Error GoTo errortrap

    Dim TOTAL_COST As Variant
    Dim LngCurItemID As Integer
    Dim LngUnitID As Long
    Dim UnitFactor As Double
    TOTAL_COST = 0
    DeleteTransactiomsVoucher val(Text1.Text)
    DeleteTransactiomsVoucher val(Text2.Text)
      
    With FG

        For i = 1 To FG.Rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("Gardresult1"))) > 0 Then
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, i, FG.ColIndex("UnitID")))
            
                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                    
                TOTAL_COST = TOTAL_COST + (Abs(FG.TextMatrix(i, FG.ColIndex("Gardresult1"))) * val(FG.TextMatrix(i, FG.ColIndex("Price"))))
            End If

        Next i

    End With

    If TOTAL_COST = 0 Then
        Exit Sub
    End If

    Dim groupAccount  As String

    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    'Dim Note_ID As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
 
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim MYTEXT As String
    '>>>>>>>>>>>>>>>>>>>>>>>>>

    rs.Close

    ' ”‰œ Ã—œ ÃœÌœ
    rs.Open "select * from Transactions where Transaction_ID = " & XPTxtBillID.Text & " and Transaction_type = 30"

    Dim xyeas As Boolean
    xyeas = True

    If xyeas = True Then
 
        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=15"))
        'mytext = TxtTransSerial.text

        '         rs!nots = mytext
        '         rs.update

        Dim Transaction_ID As Long
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Text1.Text = Transaction_ID
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        Dim TxtNoteSerialV As String
        Dim TxtNoteSerial1V As String
            
        my_branch = Me.dcBranch.BoundText

        If TxtNoteSerialV = "" Then
            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
                End If
            End If
        End If
        
        If TxtNoteSerial1V = "" Then
            If Voucher_coding(val(my_branch), XPDtbBill.value, 13, 210) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ   ”ÊÌÂ »«·⁄Ã“ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbBill.value, 13, 210) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 13, 210)
                End If
            End If
        End If
           
        Dim sql As String

        sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,NoteSerial,NoteSerial1,NoteId,BranchId,Closed)SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 15,CusID,StoreID,UserID,Emp_ID,nots=" & val(XPTxtBillID.Text) & " ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId,1From Transactions Where Transaction_ID =" & val(XPTxtBillID.Text) & " And Transaction_Type = 30"
        Cn.Execute sql
        '
        'Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID)" & "        SELECT showPrice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , abs(Gardresult)*QtyBySmalltUnit, price/QtyBySmalltUnit ,ColorID,ItemSize, UnitId,  abs(Gardresult), QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text & " and    Gardresult1>0"
          Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID)" & "        SELECT showPrice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial ,  abs(Gardresult)*QtyBySmalltUnit,Price  ,ColorID,ItemSize, UnitId,  abs(Gardresult), QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.Text & " and    Gardresult1>0"
        Text1.Text = Transaction_ID
        'TxtIssueSerial.text = TxtNoteSerial1V

        'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
        RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
        If Me.TxtModFlg.Text = "N" Then
    
        Else
        
            general_noteid = val(TxtNoteID.Text)
        End If

        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        '     TxtNoteID.text = general_noteid
        RsNotesGeneral("Transaction_ID").value = Transaction_ID
        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        RsNotesGeneral("NoteType").value = 210
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(13) '«–‰ wvt
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

        CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.dcBranch.BoundText)

    End If

End Sub

Function CREATE_VOUCHER_GE(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim TOTAL_COST As Variant
    Dim LngCurItemID As Integer
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    With FG

        For i = 1 To FG.Rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("Gardresult"))) > 0 Then
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, i, FG.ColIndex("UnitID")))
            
                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                    
                TOTAL_COST = TOTAL_COST + (Abs(FG.TextMatrix(i, FG.ColIndex("Gardresult"))) * val(FG.TextMatrix(i, FG.ColIndex("Price"))))
            End If

        Next i

    End With

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·œ«∆‰
    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)
    my_branch = BranchID

    If TOTAL_COST > 0 Then
   
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = "”‰œ  ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & Me.TxtTransSerial.Text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ      ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & TxtNoteSerial1V
            Else
                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.Rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ      ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & TxtNoteSerial1V
                        Else
                            StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·„œÌ‰
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)

        If TOTAL_COST > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

                Account_Code_dynamic = get_account_code_branch(11, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·›—Êﬁ«  «·Ã—œÌÂ   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄« 
                StrTempAccountCode = DCAccount1.BoundText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ      ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & TxtNoteSerial1V
                Else
                    StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
         
            ElseIf detect_inventory_work_type = 3 Then

                With FG

                    For i = 1 To FG.Rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 11)

                            '  groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»    ﬂ·›… «·„»Ì⁄«    ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "”‰œ      ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & TxtNoteSerial1V
                            Else
                                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                            End If
            
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If
    End If

    Dim StrSQL  As String
    StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.Text)
    Cn.Execute StrSQL
ErrTrap:
End Function

Private Sub PlusVoucher()
    ' On Error GoTo errortrap

    Dim groupAccount  As String

    DeleteTransactiomsVoucher val(Text2.Text)
   
    Dim TOTAL_COST As Variant
    Dim LngCurItemID As Integer
    Dim LngUnitID As Long
    Dim UnitFactor As Double
    TOTAL_COST = 0

    With FG

        For i = 1 To FG.Rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("Gardresult2"))) > 0 Then
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, i, FG.ColIndex("UnitID")))
            
                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                    
                TOTAL_COST = TOTAL_COST + (Abs(FG.TextMatrix(i, FG.ColIndex("Gardresult2"))) * val(FG.TextMatrix(i, FG.ColIndex("Price"))))
            End If

        Next i

    End With

    If TOTAL_COST = 0 Then
        Exit Sub
    End If

    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    'Dim Note_ID As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
 
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim MYTEXT As String
    '>>>>>>>>>>>>>>>>>>>>>>>>>

    rs.Close

    ' ”‰œ Ã—œ ÃœÌœ
    rs.Open "select * from Transactions where Transaction_ID = " & XPTxtBillID.Text & " and Transaction_type = 30"

    Dim xyeas As Boolean
    xyeas = True

    If xyeas = True Then
 
        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=15"))
        'mytext = TxtTransSerial.text

        '         rs!nots = mytext
        '         rs.update

        Dim Transaction_ID As Long
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Text2.Text = Transaction_ID
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        Dim TxtNoteSerialV As String
        Dim TxtNoteSerial1V As String
            
        my_branch = Me.dcBranch.BoundText

        If TxtNoteSerialV = "" Then
            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
                End If
            End If
        End If
        
        If TxtNoteSerial1V = "" Then
            If Voucher_coding(val(my_branch), XPDtbBill.value, 13, 210) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ   ”ÊÌÂ »«·“Ì«œ… ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbBill.value, 13, 210) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 13, 210)
                End If
            End If
        End If
           
        Dim sql As String

        sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,NoteSerial,NoteSerial1,NoteId,BranchId,Closed)SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 16,CusID,StoreID,UserID,Emp_ID,nots=" & val(XPTxtBillID.Text) & " ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId,1From Transactions Where Transaction_ID =" & val(XPTxtBillID.Text) & " And Transaction_Type = 30"
        Cn.Execute sql
        '
'        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID)" & "        SELECT showPrice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , abs(Gardresult)*QtyBySmalltUnit, price/QtyBySmalltUnit ,ColorID,ItemSize, UnitId,  abs(Gardresult), QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text & " and    Gardresult2>0"
        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID)" & "        SELECT showPrice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial ,  abs(Gardresult)*QtyBySmalltUnit,Price  ,ColorID,ItemSize, UnitId,  abs(Gardresult), QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.Text & " and    Gardresult2>0"
        
        Text2.Text = Transaction_ID
        'TxtIssueSerial.text = TxtNoteSerial1V

        'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
        RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
        If Me.TxtModFlg.Text = "N" Then
    
        Else
        
            general_noteid = val(TxtNoteID.Text)
        End If

        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        '     TxtNoteID.text = general_noteid
        RsNotesGeneral("Transaction_ID").value = Transaction_ID
        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        RsNotesGeneral("NoteType").value = 210
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(13) '«–‰ wvt
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

        CREATE_VOUCHER_GE1 Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.dcBranch.BoundText)

    End If

End Sub

Function CREATE_VOUCHER_GE1(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim TOTAL_COST As Variant
    Dim LngCurItemID As Integer
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    With FG

        For i = 1 To FG.Rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("Gardresult"))) < 0 Then
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, i, FG.ColIndex("UnitID")))
            
                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                    
                TOTAL_COST = TOTAL_COST + (Abs(FG.TextMatrix(i, FG.ColIndex("Gardresult"))) * val(FG.TextMatrix(i, FG.ColIndex("Price"))))
            End If

        Next i

    End With

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·œ«∆‰
    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)
    my_branch = BranchID

    If TOTAL_COST > 0 Then
   
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = "”‰œ  ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & Me.TxtTransSerial.Text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ      ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & TxtNoteSerial1V
            Else
                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.Rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ      ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & TxtNoteSerial1V
                        Else
                            StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·„œÌ‰
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)

        If TOTAL_COST > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

                Account_Code_dynamic = get_account_code_branch(11, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·›—Êﬁ«  «·Ã—œÌÂ   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

                StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄« 
                StrTempAccountCode = DCAccount1.BoundText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ      ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & TxtNoteSerial1V
                Else
                    StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
         
            ElseIf detect_inventory_work_type = 3 Then

                With FG

                    For i = 1 To FG.Rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 11)

                            '  groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»    ﬂ·›… «·„»Ì⁄«    ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "”‰œ      ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & TxtNoteSerial1V
                            Else
                                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                            End If
            
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If
    End If

    Dim StrSQL  As String
    StrSQL = "UPDATE Transactions SET NOTS2=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.Text)
    Cn.Execute StrSQL
ErrTrap:
End Function

Private Sub CMDStartSetelment_Click()
    'If chkStartSetelment.value = vbChecked Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    ''    MsgBox " „ ⁄„· «· ”ÊÌÂ  „‰ ﬁ»· Ê·« Ì„ﬂ‰ «· ⁄œÌ·"
    '   Else
    '   MsgBox "Cant Do"
    '   End If
    'End If
    Dim Account_Code_dynamic As String

    If DCAccount1.Text = "" Then
        Account_Code_dynamic = get_store_Account(val(DCboStoreName.BoundText), "Account_Code2")

        If Account_Code_dynamic = "" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    «· ”ÊÌ«  «·Ã—œÌÂ   ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
        
            Exit Sub
        Else
            DCAccount1.BoundText = Account_Code_dynamic
        End If
    End If
 
    MinusVoucher
    PlusVoucher
    Cmd_Click (1)
    Cmd_Click (2)
End Sub

Private Sub Command2_Click()
    Dim Transaction_ID As Integer
    Transaction_ID = val(Me.Text1.Text)

    If Transaction_ID = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmStockSettlement.show
    FrmStockSettlement.Retrive (Transaction_ID)
 
End Sub

Private Sub Command3_Click()
    Dim Transaction_ID As Integer
    Transaction_ID = val(Me.Text2.Text)

    If Transaction_ID = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmStockSettlement.show
    FrmStockSettlement.Retrive (Transaction_ID)

End Sub

Private Sub DcAccount1_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 188
            
    End If

End Sub

Private Sub DCAccount2_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 189
            
    End If

End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 12
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboStoreName_Change()

    If val(DCboStoreName.BoundText) <> 0 Then
        dcBranch.BoundText = GetInventoryBranch(DCboStoreName.BoundText)
    End If
 
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic2 As String

    Account_Code_dynamic = get_store_Account(val(DCboStoreName.BoundText), "Account_Code2")

    If Account_Code_dynamic = "" Then
        MsgBox "·„ Ì „  ÕœÌœ Õ”«»    «· ”ÊÌ«  «·Ã—œÌÂ   ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
        
        Exit Sub
    End If
        
    If Me.DCAccount1.BoundText = "" Then
        Me.DCAccount1.BoundText = Account_Code_dynamic
    End If

    If Me.DCAccount2.BoundText = "" Then
        Me.DCAccount2.BoundText = Account_Code_dynamic
    End If

End Sub

Private Sub DCboStoreName_Click(Area As Integer)
    On Error Resume Next

    DCboStoreName_Change

End Sub

Private Sub Ele_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Index = 2 Then
        If Me.WindowState = vbNormal Then
            Me.WindowState = vbMaximized
        Else
            Me.WindowState = vbNormal
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
 
    With FG

        Select Case .ColKey(Col)
            
            Case "Account_Serial"
          
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    If BolEditOnMainAccounts = False Then
 
                    End If

                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    
                    End If
 
xx:
                Else
                    GetMsgs 130, vbExclamation
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    Exit Sub
                End If

                rs.Close
                Set rs = Nothing

            Case "AccountName"
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)

                If LngRow <> -1 Then
     
                End If

                Set ClsAcc = New ClsAccounts

                If BolEditOnMainAccounts = False Then

                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
              
                Else
                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
 
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                End If

                Set ClsAcc = Nothing

        End Select

        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

    End With

End Sub

Private Sub fg_StartEdit(ByVal Row As Long, _
                         ByVal Col As Long, _
                         Cancel As Boolean)
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"

    With FG

        Select Case .ColKey(Col)

            Case "AccountName"
        
                FG.Editable = flexEDKbdMouse
                
                'Full Path Display
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '   If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '   End If
                    If True = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                    End If
                
                Else
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '     If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '     End If
                    If True = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                    End If
                
                End If
                
                Dim rs As New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = FG.BuildComboList(rs, "FirstName,ParentName,*FirstName", "Account_Code")
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub Form_Activate()
    'XPTxtBillID.SetFocus
End Sub

Private Sub Form_Load()
    Dim RsItems As New ADODB.Recordset
    Dim StrSQL As String
    Dim BGround As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Dim Msg As String

    On Error GoTo ErrTrap
    Dim My_SQL As String
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
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    Resize_Form Me, TransactionSize

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    FG.WallPaper = BGround.Picture
    AddTip
    SetDtpickerDate XPDtbBill
    NewGrid.GridTrans = NewGard
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.Grid = FG
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.GrdTBar = Me.TBr
    ' Set NewGrid.LblTotalQty = Me.LblTotalQty
    Set NewGrid.txtTotal = Me.XPTxtSum
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    NewGrid.FillGrid
    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetStores Me.DCboStoreName
 
    If SystemOptions.UserInterface = ArabicInterface Then
        Dcombos.GetAccountingCodes DCAccount1, True
        Dcombos.GetAccountingCodes DCAccount2, True
    Else
 
        Dcombos.GetAccountingCodesENg DCAccount1, True
        Dcombos.GetAccountingCodesENg DCAccount2, True

    End If

    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboStoreName

    StrSQL = "Select * From Transactions where Transaction_Type=30"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    XPBtnMove_Click 2

    TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
    Msg = Err.description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRight, App.title
End Sub

Private Sub WriteDev()
    On Error GoTo errortrap
    Dim Account_Code_dynamic As String

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then

        Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

        If Account_Code_dynamic = "" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
            Me.DcboDebitSide.BoundText = ""
            Exit Sub
        End If
        
        Me.DcboDebitSide.BoundText = Account_Code_dynamic 'Õ”«» «·„Œ“Ê‰
        'Me.DcboDebitSide.BoundText = "a1a2a5"'
    
        Account_Code_dynamic = get_account_code_branch(19, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Exit Sub
        Else

            If Account_Code_dynamic = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ   Õ”«» Ê”Ìÿ «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Exit Sub
         
            End If
        End If
        
        Me.DcboCreditSide.BoundText = Account_Code_dynamic 'Ã”«» Ê”Ìÿ «›  «ÕÌ
        'Me.DcboCreditSide.BoundText = "a2a1a1" '
 
    End If

errortrap:
End Sub

Public Function retrive1(Optional StoreId As Integer, _
                         Optional FromDate As Date, _
                         Optional ToDate As Date)
    Dim StrSQL As String
    Dim RsDetails As New ADODB.Recordset
    Dim RowNum As Integer
    Dim LngNoteID As Long

    On Error GoTo ErrTrap
 
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    StrSQL = "SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.Transaction_Details.Item_ID AS ItemID, "
    StrSQL = StrSQL & "  dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.Transaction_Details.UnitId, dbo.TblItems.ItemCode,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemName, dbo.TblUnites.UnitName, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ClassId,"
    StrSQL = StrSQL & "  dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblItemsSizes.SizeName AS SizeName, dbo.TblItemsColors.ColorName"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL + " where dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & ""
    StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & ""
    StrSQL = StrSQL & "   AND (dbo.Transactions.StoreID =" & StoreId & ")"
    StrSQL = StrSQL & "  GROUP BY dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.Transaction_Details.order_no, dbo.Transaction_Details.UnitId,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblUnites.UnitName, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize,"
    StrSQL = StrSQL & "   dbo.Transaction_Details.ClassId , dbo.TblItemsclasses.SizeName, dbo.TblItemsSizes.SizeName, dbo.TblItemsColors.ColorName"
    StrSQL = StrSQL & "  Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) <> 0)"
    StrSQL = StrSQL & "  ORDER BY dbo.TblItems.ItemID"

    Dim LngItemID As Long
    Dim LngUnitID As Long
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For RowNum = 1 To RsDetails.RecordCount

            With FG
                .TextMatrix(RowNum, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID").value), "", RsDetails("ItemID").value)
                .TextMatrix(RowNum, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID").value), "", RsDetails("ItemID").value)
                '          .TextMatrix(RowNum, FG.ColIndex("Count1")) = IIf(IsNull(RsDetails("SUMQTY").value), "", RsDetails("SUMQTY").value)
   
                LngItemID = val(.TextMatrix(RowNum, .ColIndex("Code")))
                LngUnitID = val(.TextMatrix(RowNum, .ColIndex("UnitId")))
                  
                .TextMatrix(RowNum, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , , , LngUnitID)
                    
                '.TextMatrix(RowNum, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price").value), "", RsDetails("Price").value)
                '        .TextMatrix(RowNum, FG.ColIndex("Valu")) = Val(.TextMatrix(RowNum, .ColIndex("Price"))) * Val(.TextMatrix(RowNum, .ColIndex("Count1")))
            End With

            FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), 1, (RsDetails("UnitID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
      
            RsDetails.MoveNext

            If FG.Rows > 10 Then
                If RowNum = 8 Then FG.Refresh
            End If

        Next RowNum

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    Me.XPTxtSum.Text = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Valu"), FG.Rows - 1, FG.ColIndex("Valu"))
    Me.LblTotalQty = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Count"), FG.Rows - 1, FG.ColIndex("Count"))

    NewGrid.CountItems
    Screen.MousePointer = vbDefault
    Exit Function
ErrTrap:
    Screen.MousePointer = vbDefault
End Function

Private Sub opt2_Click(Index As Integer)
    ReportType = Index
 
End Sub

Private Sub Opt3_Click(Index As Integer)
    priceType = Index
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
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

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim intDef As Integer

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.Text = "N"
            XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            Me.TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=30"))
            txtopening_balance_voucher_id.Text = get_opening_balance_voucher_id
            WriteDev
            GridDefaultValue FG.Rows - 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
            FG.SetFocus
            FG.Rows = 2
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.Rows - 1
            Me.dcBranch.BoundText = branch_id
            Dim FirstPeriodDateInthisYear  As Date
            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
            Me.DTPickerAccFrom = FirstPeriodDateInthisYear
            DTPickerAccTo.value = Date
            OPT(2).value = True

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            'If AvailableDeal = True Then
            TxtModFlg.Text = "E"
            DCboStoreName_Change
            Me.DCboUserName.BoundText = user_id

            'End If
        Case 2
            SaveData

        Case 3
            Call Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_TransAction

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
        
        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmBalanceSearch.show vbModal

        Case 6
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            '       Me.Caption = "«·—’Ìœ «·«›  «ÕÌ"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            'XPBtnAdd.Enabled = False
            'XPBtnRemove.Enabled = False
            Me.XPDtbBill.Enabled = False
            Me.DCboStoreName.locked = True
            FG.Editable = flexEDNone

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            End If

            'XPFillData.Enabled = False
            Ele(4).Enabled = False

        Case "N"
            '       Me.Caption = "«·—’Ìœ «·«›  «ÕÌ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '  Me.XPBtnMove(0).Enabled = False
            '  Me.XPBtnMove(1).Enabled = False
            '  Me.XPBtnMove(2).Enabled = False
            '  Me.XPBtnMove(3).Enabled = False
        
            Me.XPDtbBill.Enabled = True
            Me.DCboStoreName.locked = False
            XPDtbBill.value = Date
            FG.Editable = flexEDKbdMouse
        
            Ele(4).Enabled = True
            CboItemCase.ListIndex = 0

        Case "E"
            '       Me.Caption = "«·—’Ìœ «·«›  «ÕÌ(  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
                
            Me.XPDtbBill.Enabled = True
            FG.Enabled = True
            Me.DCboStoreName.locked = False
            FG.Editable = flexEDKbdMouse
            Ele(4).Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    On Error GoTo ErrTrap

    If XPTxtBillID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBillID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If AvailableDeal = True Then
                If Not rs.RecordCount < 1 Then
                    rs.delete
       
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " „ «·Õ–›"
                    Else
                        MsgBox "Deletion Done"
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
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«   ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· Â–Â «·»Ì«‰« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ Â–Â «·»Ì«‰« " & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› Â–Â «·»Ì«‰« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.Text = "R"
                XPBtnMove_Click (1)
            
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                rs.find "Transaction_ID='" & val(XPTxtBillID.Text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Exit Sub
                End If

                Me.TxtModFlg.Text = "R"
                Retrive

            End If

    End Select

    Exit Sub
ErrTrap:
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
    End If

    Set rs = Nothing
    Set cSearchDcbo = Nothing
    Set rs = Nothing
    Set TTP = Nothing
    Set BalanceReport = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim RSTransDetails As New ADODB.Recordset
    Dim RsSerial As New ADODB.Recordset
    Dim RsNotes As New ADODB.Recordset
    Dim RsCheckSerial As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSQL As String
    Dim BegineTrans As Boolean
    Dim IntRes As Integer
    Dim LngDev As Long
    Dim LngNoteID As Long
    Dim StrTemp As String

    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass
    
    If Me.TxtModFlg.Text <> "R" Then
        If DCboStoreName.BoundText = "" Then
            Msg = "ÌÃ» «Œ Ì«— «”„ «·„Œ“‰"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCboStoreName.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If NewGrid.IsReaptedSerials = True Then
            Msg = "ÌÊÃœ  ﬂ—«— ›Ï √—ﬁ«„ «·”Ì—Ì«· «·„œŒ·… "
            Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ «·√—ﬁ«„ «·„œŒ·…"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        '-----------------------------------------------
        'Check the Items Grid
        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If

        '----------------------------------------------
        Cn.BeginTrans
        BegineTrans = True

        If TxtModFlg.Text = "N" Then
            XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            Me.TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=30"))
            rs.AddNew
            rs("Transaction_ID").value = val(XPTxtBillID.Text)
        End If

     '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
  
        rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
        rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.Text)
        rs("Transaction_Serial").value = Me.TxtTransSerial.Text
        rs("Transaction_Date").value = XPDtbBill.value
    
        rs("GardFromDate").value = DTPickerAccFrom.value
        rs("GardTodate").value = DTPickerAccTo.value

        If OPT(0).value = True Then
            rs("GardEntryType").value = 0
        ElseIf OPT(1).value = True Then
            rs("GardEntryType").value = 1
        ElseIf OPT(2).value = True Then
            rs("GardEntryType").value = 2
        End If
 
        If ChkStartGard.value = vbChecked Then
            rs("StartGard").value = 1
        Else
            rs("StartGard").value = 0
        End If
 
        If chkStartSetelment.value = vbChecked Then
            rs("StartSetelment").value = 1
        Else
            rs("StartSetelment").value = 0
        End If
 
        rs("Transaction_Type").value = 30
        rs("UserID").value = user_id
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, DCboStoreName.BoundText)
        rs("Account1").value = IIf(DCAccount1.BoundText = "", Null, DCAccount1.BoundText)
        rs("Account2").value = IIf(DCAccount2.BoundText = "", Null, DCAccount2.BoundText)
    
        rs.update

        If Me.TxtModFlg.Text = "E" Then
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            ' StrSqlDel = "delete From NOTES where Transaction_ID=" & Val(rs("Transaction_ID").value)
            ' Cn.Execute StrSqlDel, , adExecuteNoRecords
    
            StrSqlDel = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.Text)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
    
        End If

        For RowNum = 1 To FG.Rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = XPTxtBillID.Text
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))

                If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                    StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                    RsCheckSerial.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsCheckSerial.EOF Or RsCheckSerial.BOF) Then
                        If RsCheckSerial("HaveSerial").value = True Then
                            RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("Serial")))
                        End If
                    End If

                    RsCheckSerial.Close
                End If

                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("Price").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            
                RSTransDetails("GardQty").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("GardQty")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("GardQty"))))
                RSTransDetails("Gardresult").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult"))))
                RSTransDetails("Gardresult1").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult1")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult1"))))
                RSTransDetails("Gardresult2").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult2")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult2"))))
            
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
                RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
                ' IIf((FG.TextMatrix(RowNum, FG.ColIndex("BranchId")) = ""), 1, Val(FG.TextMatrix(RowNum, FG.ColIndex("BranchId"))))
               
                ' RSTransDetails("ItemSize").value = _
                  IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), "", Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
                RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))

                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
        
                LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                End If

                RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                              
                RSTransDetails("showprice").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails.update
            End If

        Next RowNum
        
        Cn.CommitTrans
        BegineTrans = False
        Me.LblDevID.Caption = LngDev
        Me.LblAccountInterval.Caption = SystemOptions.SysCurrentAccountIntervalID
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
      
        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Successfully Saved " & CHR(13)
                    Msg = Msg + "Do you want to enter another  New operation"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Successfully Updated", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
            
        End Select

        TxtModFlg.Text = "R"
 
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:

    'Stop
    'Resume
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If BegineTrans = True Then
        Cn.RollbackTrans
        BegineTrans = False
    End If

    Screen.MousePointer = vbDefault

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  "
    Msg = Msg & CHR(13) & "" & Err.description
    Msg = Msg & CHR(13) & "" & Err.Number
    Msg = Msg & CHR(13) & "" & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim StrSQL As String
    Dim RsDetails As New ADODB.Recordset
    Dim RowNum As Integer
    Dim LngNoteID As Long

    On Error GoTo ErrTrap

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    Screen.MousePointer = vbArrowHourglass
    opt2(0).value = True
    Opt3(0).value = True
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", val(rs("BranchId").value))
    Text1.Text = IIf(IsNull(rs("NotS").value), "", (rs("NotS").value))
    Text2.Text = IIf(IsNull(rs("NotS2").value), "", (rs("NotS2").value))

    XPTxtBillID.Text = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
    txtopening_balance_voucher_id.Text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)
    Me.TxtTransSerial.Text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", rs("Transaction_Date").value)

    DTPickerAccFrom.value = IIf(IsNull(rs("GardFromDate").value), "", rs("GardFromDate").value)
    DTPickerAccTo.value = IIf(IsNull(rs("GardTodate").value), "", rs("GardTodate").value)

    If IsNull(rs("GardEntryType").value) Then
        OPT(2).value = True
    Else
        OPT(val(rs("GardEntryType").value)).value = True

    End If

    If IsNull(rs("StartGard").value) Then
        ChkStartGard.value = vbUnchecked
    ElseIf (rs("StartGard").value) = True Then
        ChkStartGard.value = vbChecked
    ElseIf (rs("StartGard").value) = False Then
        ChkStartGard.value = vbUnchecked

    End If

    If IsNull(rs("StartSetelment").value) Then
        chkStartSetelment.value = vbUnchecked
    ElseIf (rs("StartSetelment").value) = True Then
        chkStartSetelment.value = vbChecked
    ElseIf (rs("StartSetelment").value) = False Then
        chkStartSetelment.value = vbUnchecked

    End If

    If IsNull(rs("DifferentAccounts").value) Then
        chkDifferentAccounts.value = vbUnchecked
    ElseIf (rs("DifferentAccounts").value) = True Then
        chkDifferentAccounts.value = vbChecked
    ElseIf (rs("DifferentAccounts").value) = False Then
        chkDifferentAccounts.value = vbUnchecked

    End If

    DCAccount1.BoundText = IIf(IsNull(rs("Account1").value), "", rs("Account1").value)
    DCAccount2.BoundText = IIf(IsNull(rs("Account2").value), "", rs("Account2").value)
    DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
    'StrSql = "select * From Transaction_Details where Transaction_ID=" & Val(Rs("Transaction_ID").Value)
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Dim ItemID As Long
    Dim UnitID As Long
    Dim itemsize As Long
    Dim ColorID As Long
    Dim ClassId As Long
 
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For RowNum = 1 To RsDetails.RecordCount

            With FG
                ItemID = IIf(IsNull(RsDetails("Item_ID").value), 0, RsDetails("Item_ID").value)
                UnitID = IIf(IsNull(RsDetails("UnitID")), 1, (RsDetails("UnitID").value))
                itemsize = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
                ColorID = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
                ClassId = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
                .TextMatrix(RowNum, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID").value), "", RsDetails("Item_ID").value)
                .TextMatrix(RowNum, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID").value), "", RsDetails("Item_ID").value)
                .TextMatrix(RowNum, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Showqty").value), "", RsDetails("Showqty").value)
            
                .TextMatrix(RowNum, FG.ColIndex("GardQty")) = IIf(IsNull(RsDetails("GardQty").value), "", RsDetails("GardQty").value)
                .TextMatrix(RowNum, FG.ColIndex("Gardresult")) = IIf(IsNull(RsDetails("Gardresult").value), "", RsDetails("Gardresult").value)
                .TextMatrix(RowNum, FG.ColIndex("Gardresult1")) = IIf(IsNull(RsDetails("Gardresult1").value), "", RsDetails("Gardresult1").value)
                .TextMatrix(RowNum, FG.ColIndex("Gardresult2")) = IIf(IsNull(RsDetails("Gardresult2").value), "", RsDetails("Gardresult2").value)
            
                .TextMatrix(RowNum, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial").value), "", RsDetails("ItemSerial").value)
                .TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = IIf(IsNull(RsDetails("HaveSerial").value), "", RsDetails("HaveSerial").value)

                If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                    FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
                Else
                    FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
                End If

                .TextMatrix(RowNum, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice").value), "", RsDetails("ShowPrice").value)
                .TextMatrix(RowNum, FG.ColIndex("Valu")) = val(.TextMatrix(RowNum, .ColIndex("Price"))) * val(.TextMatrix(RowNum, .ColIndex("Count")))
            End With

            FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
      
            'FG.TextMatrix(RowNum, FG.ColIndex("GardQty")) = GetActualItemQty(Val(Me.DCboStoreName.BoundText), Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value, _
             ItemID, UnitId, ItemSize, ColorID, ClassId)
            'FG.TextMatrix(RowNum, FG.ColIndex("Gardresult")) = Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) - Val(FG.TextMatrix(RowNum, FG.ColIndex("GardQty")))
            ' «·›—ﬁ Ì”«ÊÌ „«  „ «œŒ«·Â
            '„ÿ—ÊÕ „‰Â «·„ÊÃÊœ ›⁄·« »«·»—‰«„Ã
 
            RsDetails.MoveNext

            If FG.Rows > 10 Then
                If RowNum = 8 Then FG.Refresh
            End If

        Next RowNum

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    Me.XPTxtSum.Text = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Valu"), FG.Rows - 1, FG.ColIndex("Valu"))
    Me.LblTotalQty = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Count"), FG.Rows - 1, FG.ColIndex("Count"))
 
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    NewGrid.CountItems
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub printing()
    On Error GoTo ErrTrap

    If XPTxtBillID.Text <> "" Then
        Set BalanceReport = New ClsOpeningBalanceReport
        BalanceReport.ShowOpeningBalanceData XPTxtBillID.Text
    End If

    Exit Sub
ErrTrap:
End Sub

Public Function GardReport(Transactionid As Integer, _
                           ReportType As Integer, _
                           priceType As Integer)
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim cOptions As ClsCompanyInfo
    Dim rs As New ADODB.Recordset
    Dim CViewer As ClsReportViewer
    Dim Msg As String
    Dim Reportpath As String

    'On Error GoTo ErrTrap
    Reportpath = (App.path & "\Reports\Inventory\Gard3.rpt")
    StrSQL = "SELECT     dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.Transaction_ID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.Transaction_Details.ShowQty, "
    StrSQL = StrSQL & "  dbo.Transaction_Details.showPrice, dbo.TblItemsColors.ColorName, dbo.TblItemsclasses.SizeName AS className, dbo.Transactions.Transaction_Date,"
    StrSQL = StrSQL & "  dbo.Transactions.Transaction_Type, dbo.Transaction_Details.GardQty, dbo.Transaction_Details.Gardresult, dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName,"
    StrSQL = StrSQL & "  dbo.TblStore.StoreName , dbo.TblItems.SallingPrice"
    StrSQL = StrSQL & "  FROM         dbo.Transaction_Details INNER JOIN"
    StrSQL = StrSQL & "  dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL & "  WHERE     (dbo.Transaction_Details.Transaction_ID = " & Transactionid & ") "

    If ReportType = 0 Then
 
    ElseIf ReportType = 1 Then
        StrSQL = StrSQL & " AND (dbo.Transaction_Details.Gardresult < 0)"
  
    ElseIf ReportType = 2 Then
        StrSQL = StrSQL & " AND (dbo.Transaction_Details.Gardresult > 0)"
    End If

    StrSQL = StrSQL & "  ORDER BY dbo.Transaction_Details.Item_ID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(App.path & "\Reports\Inventory\Gard3.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + Reportpath
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    Else
        Screen.MousePointer = vbArrowHourglass
        Set xReport = xApp.OpenReport(Reportpath)
        xReport.Database.SetDataSource rs
         
        Set cOptions = New ClsCompanyInfo
        xReport.ParameterFields(1).AddCurrentValue cOptions.ArabCompanyName
        xReport.ParameterFields(2).AddCurrentValue cOptions.ArabComment
        xReport.ParameterFields(3).AddCurrentValue user_name
        xReport.ParameterFields(4).AddCurrentValue ""
        xReport.ParameterFields(5).AddCurrentValue GetCurrentGardEmployee
          
        xReport.ParameterFields(6).AddCurrentValue Format(DTPickerAccTo, "yyyy/m/d")
        xReport.ParameterFields(7).AddCurrentValue Format(DTPickerAccFrom, "yyyy/m/d")

        If priceType = 0 Then
            xReport.ParameterFields(8).AddCurrentValue "costprice"
        Else
            xReport.ParameterFields(8).AddCurrentValue "SalePrice"
        End If
         
        Screen.MousePointer = vbDefault
    End If

    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, , , , , Reportpath
    Screen.MousePointer = vbDefault
    
    Exit Function
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault
End Function

Private Function AvailableDeal() As Boolean
    Dim RowNum As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsSalle As ADODB.Recordset
    Dim LngItemID As Long
    On Error GoTo ErrTrap

    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            StrSQL = "select * From QryDelPurchase where Transaction_Date>=" & SQLDate(XPDtbBill.value, True) & ""
            StrSQL = StrSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))

            '        If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) <> "" Then
            '            If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = True Then
            If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                End If

                '            End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then

                    '                StrSql = "select * From QryGardComplete where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                    '                StrSql = StrSql + " AND ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                    '                StrSql = StrSql + " AND StoreID=" & DCboStoreName.BoundText
                    '                Set RsTemp = New ADODB.Recordset
                    '                RsTemp.Open StrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    '                If RsTemp.EOF Or RsTemp.BOF Then
                    With FrmAlarm
                        .DealingForm = OpeningBalance
                        .show vbModal
                    End With

                    AvailableDeal = False
                    Exit Function
                    '                End If
                    RsTemp.Close
                Else
                    LngItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                    Set RsTemp = New ADODB.Recordset
                    Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, val(Me.XPTxtBillID.Text))

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If val(RsTemp("totalqty").value) < val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) Then

                            With FrmAlarm
                                .DealingForm = OpeningBalance
                                .show vbModal
                            End With

                            AvailableDeal = False
                            Exit Function
                        End If
                    End If

                    RsTemp.Close
                End If
            End If

            RsSalle.Close
        End If

    Next RowNum

    AvailableDeal = True
    Exit Function
ErrTrap:
End Function

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

Private Sub GridDefaultValue(Crow As Long)
    On Error GoTo ErrTrap

    With FG
        .TextMatrix(Crow, .ColIndex("ItemCase")) = 1
        .TextMatrix(Crow, .ColIndex("Count")) = 1
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            '        Cmd_Click (0)
        Else
            '        SendKeys "{TAB}"
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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If KeyCode = vbKeyF2 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
        
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
        
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            
            End If
        End If
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

Private Sub ChangeLang()
    On Error Resume Next
 
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Asset Adjustement"
    C1Elastic6.Caption = Me.Caption

    lbl(1).Caption = "ID"
    lbl(0).Caption = "Date"
    Label3.Caption = "Branch"
Check1.Caption = "All Data"
Ele(2).Caption = "Period"
lbl(10).Caption = "From"
lbl(11).Caption = "To"

    lbl(2).Caption = "Store "
    lbl(63).Caption = "Total Qty "

    Fra.Caption = "GL"
    lbl(8).Caption = "GL#"
    lbl(9).Caption = "Interval"
    lbl(32).Caption = "Depit"
    lbl(7).Caption = "Credit"

    lbl(3).Caption = " Total:"
 
    lbl(6).Caption = " By:"
    lbl(4).Caption = "Curr. rec"
    lbl(5).Caption = "Rec. Count:"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = " Case"
    lbl(30).Caption = " Master P"
    
    lbl(28).Caption = " Serial"
    lbl(27).Caption = "QTY"
    lbl(26).Caption = "Price"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
   
    lbl(12).Caption = "Sett Account"
    CMDStartGard.Caption = "Start Inv."
    Command2.Caption = "Dec Voucher"
    Command3.Caption = "Inc Voucher"
    Frame2.Caption = "Report Setting"
    opt2(0).Caption = "Print All"
    opt2(1).Caption = "Dec Only"
    opt2(2).Caption = "Inc Only"
    Frame3.Caption = "Show Price"
    Opt3(0).Caption = "Cost Price"
    Opt3(1).Caption = "Sale Price"
    CMdPrit.Caption = "Print"
   

 
    With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("code")) = "Asset Code"
        .TextMatrix(0, .ColIndex("Name")) = "Asset Name "
        .TextMatrix(0, .ColIndex("Era")) = "Employee"
        .TextMatrix(0, .ColIndex("count")) = "Actual Qty"
        
        .TextMatrix(0, .ColIndex("GardQty")) = "Qty"
        .TextMatrix(0, .ColIndex("Gardresult")) = "Different "
        .TextMatrix(0, .ColIndex("Gardresult1")) = "Deficit"
        .TextMatrix(0, .ColIndex("Gardresult2")) = "Increase"
       .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"

        
    End With

    CMDStartSetelment.Caption = "Start Settlement"

    'NewItem

End Sub

Private Sub XPTxtSum_Change()
    Me.lblTotal.Caption = XPTxtSum.Text
    Exit Sub
ErrTrap:
End Sub
