VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmNewGard1 
   Caption         =   " ‰›Ì– «·Ã—œ «·›⁄·Ï ··„Œ«“‰"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13485
   HelpContextID   =   90
   Icon            =   "FrmNewGard1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmNewGard1.frx":038A
   RightToLeft     =   -1  'True
   ScaleHeight     =   7755
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
      Height          =   7755
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   13485
      _cx             =   23786
      _cy             =   13679
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
      _GridInfo       =   $"FrmNewGard1.frx":0714
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   465
         Index           =   5
         Left            =   15
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   7275
         Width           =   13455
         _cx             =   23733
         _cy             =   820
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
            Height          =   225
            Left            =   10005
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   45
            Visible         =   0   'False
            Width           =   270
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   2400
            TabIndex        =   12
            Top             =   45
            Width           =   1320
            _ExtentX        =   2328
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
            Height          =   195
            Index           =   63
            Left            =   5430
            TabIndex        =   62
            Top             =   90
            Visible         =   0   'False
            Width           =   975
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
            Height          =   240
            Left            =   4815
            TabIndex        =   61
            Top             =   0
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label lblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   9855
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·—’Ìœ"
            Height          =   150
            Index           =   3
            Left            =   10755
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   90
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   180
            Index           =   6
            Left            =   3750
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   90
            Width           =   585
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   135
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   105
            Width           =   555
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   150
            Left            =   1500
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   90
            Width           =   390
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   5
            Left            =   810
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   0
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   4
            Left            =   1980
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   0
            Width           =   390
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   5145
         Index           =   3
         Left            =   15
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2115
         Width           =   13455
         _cx             =   23733
         _cy             =   9075
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
         GridRows        =   4
         GridCols        =   6
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmNewGard1.frx":07A7
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin MSComctlLib.Toolbar TBr 
            Height          =   630
            Left            =   510
            TabIndex        =   26
            Top             =   4875
            Width           =   12420
            _ExtentX        =   21908
            _ExtentY        =   1111
            ButtonWidth     =   609
            ButtonHeight    =   1005
            Appearance      =   1
            _Version        =   393216
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   990
            Index           =   4
            Left            =   30
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   13395
            _cx             =   23627
            _cy             =   1746
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
            Begin VB.TextBox TxtShortName 
               Height          =   255
               Left            =   720
               TabIndex        =   106
               Top             =   120
               Width           =   6795
            End
            Begin VB.TextBox TxtPrice 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   675
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   615
               Width           =   1650
            End
            Begin VB.TextBox TxtSerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   4065
               MaxLength       =   20
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   615
               Width           =   2055
            End
            Begin VB.TextBox TxtQuantity 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   2445
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   615
               Width           =   1530
            End
            Begin VB.ComboBox CboItemCase 
               Height          =   315
               Left            =   6165
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   615
               Width           =   1965
            End
            Begin MSDataListLib.DataCombo DCboItemsName 
               Height          =   315
               Left            =   8130
               TabIndex        =   1
               Top             =   615
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboItemsCode 
               Height          =   315
               Left            =   10725
               TabIndex        =   0
               Top             =   615
               Width           =   2640
               _ExtentX        =   4657
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdAdd 
               Height          =   390
               Left            =   30
               TabIndex        =   6
               Top             =   510
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   688
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
               ButtonImage     =   "FrmNewGard1.frx":0830
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
               Caption         =   "«·»ÕÀ «·”—Ì⁄"
               Height          =   270
               Index           =   97
               Left            =   7950
               TabIndex        =   107
               Top             =   120
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”⁄—"
               Height          =   240
               Index           =   26
               Left            =   855
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   375
               Width           =   1470
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   240
               Index           =   27
               Left            =   2670
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   375
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ì—Ì«·"
               Height          =   360
               Index           =   28
               Left            =   4215
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   375
               Width           =   1950
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·’‰›"
               Height          =   240
               Index           =   29
               Left            =   6300
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   375
               Width           =   1830
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈”„ «·’‰›"
               Height          =   240
               Index           =   30
               Left            =   8355
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   375
               Width           =   2370
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   240
               Index           =   31
               Left            =   10950
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   375
               Width           =   2430
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   3825
            Left            =   30
            TabIndex        =   60
            Top             =   1035
            Width           =   13395
            _cx             =   23627
            _cy             =   6747
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
            Cols            =   23
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmNewGard1.frx":0BCA
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
            Editable        =   1
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
            Height          =   240
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   4875
            Width           =   450
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1425
         Index           =   1
         Left            =   15
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   675
         Width           =   13455
         _cx             =   23733
         _cy             =   2514
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
         Begin VB.CheckBox chkIsAutoCost 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "«· ﬂ·›… ¬·Ì«"
            Enabled         =   0   'False
            Height          =   225
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   420
            Width           =   1125
         End
         Begin VB.CheckBox chkAutoDetect 
            Alignment       =   1  'Right Justify
            Caption         =   "Ã—œ «·«’‰«› €Ì— «·„ÊÃÊœ… » ﬂ·›… ’›—"
            Height          =   255
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CommandButton Command3 
            Caption         =   "”‰œ«   «·“Ì«œÂ"
            Height          =   345
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   1095
            Width           =   1095
         End
         Begin VB.Frame Frame2 
            Caption         =   "ŒÌ«—«  ÿ»«⁄Â «·Ã—œ"
            Height          =   1380
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   0
            Width           =   3255
            Begin VB.OptionButton opt2 
               Alignment       =   1  'Right Justify
               Caption         =   "«·’›—Ì"
               Height          =   195
               Index           =   3
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   960
               Width           =   1335
            End
            Begin VB.CommandButton cmdSearch 
               Caption         =   "»ÕÀ"
               Height          =   315
               Left            =   1920
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   1200
               Width           =   1245
            End
            Begin VB.CommandButton CMdPrit 
               Caption         =   "ÿ»«⁄Â"
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   1200
               Width           =   1695
            End
            Begin VB.Frame Frame3 
               Caption         =   "⁄—÷ «·”⁄—"
               Height          =   735
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   120
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
               TabIndex        =   95
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
               TabIndex        =   94
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
               TabIndex        =   93
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   11880
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   -270
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "”‰œ«   «·⁄Ã“"
            Height          =   345
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1095
            Width           =   1095
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   12120
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   1095
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   435
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   135
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton CMDStartSetelment 
            Caption         =   " ‰›Ì– «· ”ÊÌ«  «·Ã—œÌÂ"
            Enabled         =   0   'False
            Height          =   345
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   720
            Width           =   2175
         End
         Begin VB.CommandButton CMDStartGard 
            Caption         =   " ‰›Ì– «·Ã—œ"
            Enabled         =   0   'False
            Height          =   330
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   360
            Width           =   2175
         End
         Begin VB.Frame Frame1 
            Caption         =   "Õœœ ÿ—Ìﬁ… «·«œŒ«·"
            Height          =   1110
            Left            =   13470
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   555
            Visible         =   0   'False
            Width           =   3315
            Begin VB.CommandButton Command1 
               Caption         =   " ÕœÌœ «·„·›..."
               Height          =   255
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   71
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
               TabIndex        =   70
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
               TabIndex        =   69
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
               TabIndex        =   68
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.TextBox txtopening_balance_voucher_id 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   2925
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   1635
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
            Height          =   1410
            Left            =   -7920
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   0
            Width           =   8010
            Begin MSDataListLib.DataCombo DcboDebitSide 
               Height          =   315
               Left            =   1680
               TabIndex        =   35
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
               TabIndex        =   36
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
               TabIndex        =   42
               Top             =   510
               Width           =   1095
            End
            Begin VB.Label LblDevID 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   5580
               RightToLeft     =   -1  'True
               TabIndex        =   41
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
               TabIndex        =   40
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
               TabIndex        =   39
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
               TabIndex        =   38
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
               TabIndex        =   37
               Top             =   180
               Width           =   885
            End
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   105
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   1170
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   405
            Left            =   11130
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   60
            Width           =   1350
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   1035
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   705
            Visible         =   0   'False
            Width           =   825
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   375
            Left            =   8400
            TabIndex        =   99
            Top             =   0
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   101842947
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   9930
            TabIndex        =   84
            Top             =   810
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   9930
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   555
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   735
            Index           =   2
            Left            =   7410
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   510
            Width           =   2475
            _cx             =   4366
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
            Enabled         =   0   'False
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   16711680
            FloodColor      =   16711680
            ForeColorDisabled=   -2147483631
            Caption         =   " «—ÌŒ «·Ã—œ"
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
               Left            =   840
               TabIndex        =   86
               ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
               Top             =   1080
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   -2147483624
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   101842947
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
               Left            =   210
               TabIndex        =   87
               ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
               Top             =   360
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   -2147483624
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   101842947
               CurrentDate     =   37357
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·Ã—œ"
               ForeColor       =   &H00FF8080&
               Height          =   285
               Index           =   11
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   360
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   285
               Index           =   10
               Left            =   2070
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   1125
               Visible         =   0   'False
               Width           =   555
            End
         End
         Begin MSDataListLib.DataCombo DCAccount1 
            Height          =   315
            Left            =   3840
            TabIndex        =   88
            Top             =   0
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
            TabIndex        =   77
            Top             =   840
            Visible         =   0   'False
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «· ”ÊÌÂ »«·“Ì«œ…"
            Height          =   420
            Index           =   13
            Left            =   5715
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   810
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «· ”ÊÌÂ "
            Height          =   420
            Index           =   12
            Left            =   6195
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   0
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   270
            Index           =   2
            Left            =   12405
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   555
            Width           =   945
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   12600
            TabIndex        =   59
            Top             =   810
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Õ—ﬂ…"
            Height          =   435
            Index           =   0
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   120
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”·”·"
            Height          =   450
            Index           =   1
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   75
            Width           =   600
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   645
         Left            =   15
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   15
         Width           =   13455
         _cx             =   23733
         _cy             =   1138
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
         Caption         =   " ‰›Ì– «·Ã—œ «·›⁄·Ï ··„Œ«“‰  "
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
         Begin VB.TextBox TxtValueAdded 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Text            =   "0"
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CheckBox chkDifferentAccounts 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "«· ”ÊÌÂ ⁄·Ï Õ”«»«  „Œ ·›…"
            Height          =   210
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   390
            Width           =   2535
         End
         Begin VB.CheckBox ChkStartGard 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   " „  ‰›Ì– «·Ã—œ"
            Enabled         =   0   'False
            Height          =   225
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   420
            Width           =   1335
         End
         Begin VB.CheckBox chkStartSetelment 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   " „  ‰›Ì– «· ”ÊÌ«  «·Ã—œÌÂ"
            Enabled         =   0   'False
            Height          =   165
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   165
            Width           =   2055
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   315
            Index           =   0
            Left            =   1755
            TabIndex        =   44
            Top             =   165
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
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
            ButtonImage     =   "FrmNewGard1.frx":0F86
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
            Height          =   315
            Index           =   3
            Left            =   975
            TabIndex        =   45
            Top             =   165
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   556
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
            ButtonImage     =   "FrmNewGard1.frx":1320
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
            Height          =   315
            Index           =   1
            Left            =   2700
            TabIndex        =   46
            Top             =   165
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
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
            ButtonImage     =   "FrmNewGard1.frx":16BA
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
            Height          =   315
            Index           =   2
            Left            =   150
            TabIndex        =   47
            Top             =   165
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
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
            ButtonImage     =   "FrmNewGard1.frx":1A54
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin VSFlex8UCtl.VSFlexGrid VatGrid 
            Height          =   405
            Left            =   1920
            TabIndex        =   102
            Tag             =   "1"
            Top             =   0
            Visible         =   0   'False
            Width           =   3720
            _cx             =   6562
            _cy             =   714
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
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmNewGard1.frx":1DEE
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ã«—Ì  ‰›Ì– «·⁄„·Ì… »—Ã«¡ «·«‰ Ÿ«—..."
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   165
            Visible         =   0   'False
            Width           =   5055
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   900
         Index           =   0
         Left            =   15
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   6840
         Visible         =   0   'False
         Width           =   13455
         _cx             =   23733
         _cy             =   1588
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
            Height          =   810
            Index           =   0
            Left            =   18120
            TabIndex        =   49
            Top             =   285
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   1429
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
            Height          =   810
            Index           =   1
            Left            =   15915
            TabIndex        =   50
            Top             =   285
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   1429
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
            Height          =   795
            Index           =   2
            Left            =   13620
            TabIndex        =   51
            Top             =   345
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   1402
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
            Height          =   810
            Index           =   3
            Left            =   11535
            TabIndex        =   52
            Top             =   285
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   1429
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
            Height          =   810
            Index           =   4
            Left            =   8700
            TabIndex        =   53
            Top             =   285
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   1429
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
            Height          =   810
            Index           =   5
            Left            =   6825
            TabIndex        =   54
            Top             =   285
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   1429
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
            Height          =   810
            Index           =   6
            Left            =   45
            TabIndex        =   55
            Top             =   285
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1429
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
            Height          =   810
            Index           =   7
            Left            =   4395
            TabIndex        =   56
            Top             =   285
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   1429
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
            Height          =   810
            Left            =   2430
            TabIndex        =   57
            Top             =   285
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   1429
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
Attribute VB_Name = "FrmNewGard1"
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

Private Sub Check1_Click()

End Sub

Private Sub chkDifferentAccounts_Click()

    If chkDifferentAccounts.value = vbChecked Then
        FG.Editable = flexEDKbdMouse
    Else
        FG.Editable = flexEDNone

    End If

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CMdPrit_Click()
    GardReport val(Me.XPTxtBillID.text), ReportType, priceType
End Sub

Private Sub CmdSearch_Click()
    If DoPremis(Do_Search, Me.Name, True) = False Then
        Exit Sub
    End If
    FrmBalanceSearch.mTransaction_Type = 30
    FrmBalanceSearch.mIndex = 1
    FrmBalanceSearch.show vbModal
End Sub

Private Sub CMDStartGard_Click()
                             If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                      
                      
    'If ChkStartGard.value = vbChecked Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '    MsgBox " „ ⁄„· «·Ã—œ „‰ ﬁ»· Ê·« Ì„ﬂ‰ «· ⁄œÌ·"
    '    Else
    '    MsgBox "Cant Do"
    '    End If
    '    Exit Sub
    'End If
    Label1.Visible = True
 DoEvents
    DeleteTransactiomsVoucher val(Text1.text)
    DeleteTransactiomsVoucher val(Text2.text)
    Text1.text = ""
    Text2.text = ""
    Dim ItemID As Long
    Dim UnitID As Long
    Dim itemsize As Long
    Dim ColorID As Long
    Dim ClassId As Long
    Dim ParrtNoCode As String

    With FG

        For RowNum = 1 To FG.rows - 1
        
            ItemID = val(.TextMatrix(RowNum, FG.ColIndex("Code")))
            UnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            itemsize = val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")))
            ColorID = val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID")))
            ClassId = val(FG.TextMatrix(RowNum, FG.ColIndex("ClassID")))
      
'            Fg.TextMatrix(RowNum, Fg.ColIndex("GardQty")) = GetActualItemQty(val(Me.DCboStoreName.BoundText), Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value, ItemID, unitid, itemsize, ColorID, ClassId)
 Dim FirstPeriodDateInthisYear  As Date
  
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            Me.DTPickerAccFrom = FirstPeriodDateInthisYear
            
            ParrtNoCode = (FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")))
            
   
                  
            
                If ParrtNoCode = "" Then
                
 
                                If FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) <> "" And IsDate(FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))) Then
                                           FG.TextMatrix(RowNum, FG.ColIndex("GardQty")) = GetActualItemQtyNew(val(Me.DCboStoreName.BoundText), Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value, ItemID, UnitID, itemsize, ColorID, ClassId, FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")), 4)
                                 Else
                                           FG.TextMatrix(RowNum, FG.ColIndex("GardQty")) = GetActualItemQtyNew(val(Me.DCboStoreName.BoundText), Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value, ItemID, UnitID, itemsize, ColorID, ClassId, , 4)
                                 End If
              
                Else
                FG.TextMatrix(RowNum, FG.ColIndex("GardQty")) = GetQtyByBarcode(val(Me.DCboStoreName.BoundText), Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value, ParrtNoCode)
                
                End If
           FG.TextMatrix(RowNum, FG.ColIndex("Gardresult")) = val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) - val(FG.TextMatrix(RowNum, FG.ColIndex("GardQty")))
           
           
           
               If val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult"))) <> 0 And val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))) <= 0 Then
                         
               End If
               If chkIsAutoCost.value = vbChecked Then
                      '.TextMatrix(RowNum, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(CLng(ItemID), 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, 0, UnitID)
                      .TextMatrix(RowNum, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPriceByGard(CLng(ItemID), 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.XPTxtBillID), UnitID)
                End If
               
               If val(FG.TextMatrix(RowNum, FG.ColIndex("AutoDetect"))) = 1 And chkAutoDetect.value = vbChecked Then
               .TextMatrix(RowNum, FG.ColIndex("Price")) = 0
               End If
               
            If val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult"))) < 0 Then
                      
   
                FG.TextMatrix(RowNum, FG.ColIndex("Gardresult2")) = Abs(val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult"))))
                FG.TextMatrix(RowNum, FG.ColIndex("Gardresult1")) = 0
                
                
            ElseIf val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult"))) >= 0 Then
                      
              
          

                FG.TextMatrix(RowNum, FG.ColIndex("Gardresult1")) = (Abs(val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult")))))
                FG.TextMatrix(RowNum, FG.ColIndex("Gardresult2")) = 0
                Else
                
                
            End If
 
        Next RowNum

        FG.AutoSize 0, FG.Cols - 1, False
    End With

    ChkStartGard.value = vbChecked
 
    Cmd_Click (1)
    Cmd_Click (2)
    Label1.Visible = False
End Sub

Private Sub MinusVoucher()
    ' On Error GoTo errortrap

    Dim TOTAL_COST As Variant
    Dim LngCurItemID As Long
    Dim LngCurItemID2 As Long
    Dim LngUnitID As Long
    Dim UnitFactor As Double
    TOTAL_COST = 0
   
    DeleteTransactiomsVoucher val(Text1.text)
    'DeleteTransactiomsVoucher val(Text2.text)
      
    With FG

        For i = 1 To FG.rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("Gardresult1"))) > 0 Then
                LngCurItemID2 = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, i, FG.ColIndex("UnitID")))
            
                GetUnitNoOfItems LngCurItemID2, LngUnitID, UnitFactor
                    
                TOTAL_COST = TOTAL_COST + (Abs(FG.TextMatrix(i, FG.ColIndex("Gardresult1"))) * val(FG.TextMatrix(i, FG.ColIndex("Price"))))
            End If

        Next i

    End With

    'If TOTAL_COST = 0 Then
    '    Exit Sub
    'End If

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
    rs.Open "select * from Transactions where Transaction_ID = " & XPTxtBillID.text & " and Transaction_type = 30"

    Dim xyeas As Boolean
    xyeas = True

    If xyeas = True Then
 
        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=15"))
        'mytext = TxtTransSerial.text

        '         rs!nots = mytext
        '         rs.update

        Dim Transaction_ID As Long
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Text1.text = Transaction_ID
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

        sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,NoteSerial,NoteSerial1,NoteId,BranchId,Closed)SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 15,CusID,StoreID,UserID,Emp_ID,nots=" & val(XPTxtBillID.text) & " ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId,1From Transactions Where Transaction_ID =" & val(XPTxtBillID.text) & " And Transaction_Type = 30"
        Cn.Execute sql
        '
        'Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID)" & "        SELECT showPrice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , abs(Gardresult)*QtyBySmalltUnit, price/QtyBySmalltUnit ,ColorID,ItemSize, UnitId,  abs(Gardresult), QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text & " and    Gardresult1>0"
          Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ExpiryDate)" & "        SELECT showPrice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial ,  abs(Gardresult)*QtyBySmalltUnit,Price  ,ColorID,ItemSize, UnitId,  abs(Gardresult), QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ExpiryDate  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text & " and    Gardresult1>0"
          Cn.Execute "INSERT INTO  dbo.ItemsDetails(Transaction_ID,ItemId,ItemDetailedCode,ParrtNoCode,UnitID,ColorID,SizeID,ClassId,Count,EffectN )" & "  select      " & Transaction_ID & ",Item_ID,ItemDetailedCode,ParrtNoCode,UnitID,ColorID,ItemSize,ClassId,abs(Gardresult) ,1 From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text & " and    Gardresult1>0"
          
             UpdateTransactionsCost CStr(Transaction_ID)
        Text1.text = Transaction_ID
        'TxtIssueSerial.text = TxtNoteSerial1V

        'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
        'RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
        If Me.TxtModFlg.text = "N" Then
    
        Else
        
            general_noteid = val(TXTNoteID.text)
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
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    With FG

        For i = 1 To FG.rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("Gardresult"))) > 0 Then
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, i, FG.ColIndex("UnitID")))
            
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
            StrTempDes = "”‰œ  ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & Me.TxtTransSerial.text
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

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.rows - 1

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

'                Account_Code_dynamic = get_account_code_branch(11, my_branch)
        
'                If Account_Code_dynamic = "NO branch" Then
'                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'                    GoTo ErrTrap
'                Else
'
'                    If Account_Code_dynamic = "NO account" Then
'                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·›—Êﬁ«  «·Ã—œÌÂ   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                        GoTo ErrTrap
'
'                    End If
'                End If
'
'                StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄« 
                StrTempAccountCode = DcAccount1.BoundText

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

                With FG

                    For i = 1 To FG.rows - 1

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
    StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
    Cn.Execute StrSQL
ErrTrap:
End Function

Private Sub PlusVoucher()
    ' On Error GoTo errortrap

    Dim groupAccount  As String
 If val(Text1.text) = val(Text2.text) And ((Text1.text) <> "") Then Exit Sub
    DeleteTransactiomsVoucher val(Text2.text)
   
    Dim TOTAL_COST As Variant
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim UnitFactor As Double
    TOTAL_COST = 0

    With FG

        For i = 1 To FG.rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("Gardresult2"))) > 0 Then
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, i, FG.ColIndex("UnitID")))
            
                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                    
                TOTAL_COST = TOTAL_COST + (Abs(FG.TextMatrix(i, FG.ColIndex("Gardresult2"))) * val(FG.TextMatrix(i, FG.ColIndex("Price"))))
            End If

        Next i

    End With

 '   If TOTAL_COST = 0 Then
 '       Exit Sub
 '   End If

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
    rs.Open "select * from Transactions where Transaction_ID = " & XPTxtBillID.text & " and Transaction_type = 30"

    Dim xyeas As Boolean
    xyeas = True

    If xyeas = True Then
 
        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=15"))
        'mytext = TxtTransSerial.text

        '         rs!nots = mytext
        '         rs.update

        Dim Transaction_ID As Long
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Text2.text = Transaction_ID
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

        sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,NoteSerial,NoteSerial1,NoteId,BranchId,Closed)SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 16,CusID,StoreID,UserID,Emp_ID,nots=" & val(XPTxtBillID.text) & " ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId,1From Transactions Where Transaction_ID =" & val(XPTxtBillID.text) & " And Transaction_Type = 30"
        Cn.Execute sql
        '
'        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID)" & "        SELECT showPrice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , abs(Gardresult)*QtyBySmalltUnit, price/QtyBySmalltUnit ,ColorID,ItemSize, UnitId,  abs(Gardresult), QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text & " and    Gardresult2>0"
        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ExpiryDate)" & "        SELECT showPrice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial ,  abs(Gardresult)*QtyBySmalltUnit,Price  ,ColorID,ItemSize, UnitId,  abs(Gardresult), QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ExpiryDate  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text & " and    Gardresult2>0"
              Cn.Execute "INSERT INTO  dbo.ItemsDetails(Transaction_ID,ItemId,ItemDetailedCode,ParrtNoCode,UnitID,ColorID,SizeID,ClassId,Count,EffectN )" & "  select      " & Transaction_ID & ",Item_ID,ItemDetailedCode,ParrtNoCode,UnitID,ColorID,ItemSize,ClassId,abs(Gardresult) ,-1 From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text & " and    Gardresult2>0"
                 UpdateTransactionsCost CStr(Transaction_ID)
                 
        Text2.text = Transaction_ID
        'TxtIssueSerial.text = TxtNoteSerial1V

        'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
     '   RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
        If Me.TxtModFlg.text = "N" Then
    
        Else
        
            general_noteid = val(TXTNoteID.text)
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
    Dim LngCurItemID As Long
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    With FG

        For i = 1 To FG.rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("Gardresult"))) < 0 Then
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, i, FG.ColIndex("UnitID")))
            
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
            StrTempDes = "”‰œ  ”ÊÌ«  Ã—œÌÂ  —ﬁ„ " & Me.TxtTransSerial.text
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

                For i = 1 To FG.rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
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

     '           Account_Code_dynamic = get_account_code_branch(11, my_branch)
        
     '           If Account_Code_dynamic = "NO branch" Then
     '               MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
     '               GoTo ErrTrap
     '           Else
'
'                    If Account_Code_dynamic = "NO account" Then
'                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·›—Êﬁ«  «·Ã—œÌÂ   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                        GoTo ErrTrap
'
'                    End If
'                End If

                StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄« 
                StrTempAccountCode = DcAccount1.BoundText

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

                    For i = 1 To FG.rows - 1

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
    StrSQL = "UPDATE Transactions SET NOTS2=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
    Cn.Execute StrSQL
ErrTrap:
End Function

Private Sub CMDStartSetelment_Click()
                             If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                      
                      
 
    Label1.Visible = True
    Dim Account_Code_dynamic As String

    If DcAccount1.text = "" Then
        Account_Code_dynamic = get_store_Account(val(DCboStoreName.BoundText), "Account_Code2")

        If Account_Code_dynamic = "" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»    «· ”ÊÌ«  «·Ã—œÌÂ   ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
        
            Exit Sub
        Else
            DcAccount1.BoundText = Account_Code_dynamic
        End If
    End If
 CMDStartGard_Click
 
    MinusVoucher
    PlusVoucher
    Cmd_Click (1)
    Cmd_Click (2)
    Label1.Visible = False
End Sub

Private Sub Command2_Click()
    Dim Transaction_ID As Double
    Transaction_ID = val(Me.Text1.text)

    If Transaction_ID = 0 Then MsgBox "€Ì— „”Ã· Â–« «·”‰œ": Exit Sub
    FrmStockSettlement.show
    FrmStockSettlement.Retrive (Transaction_ID)
 
End Sub

Private Sub Command3_Click()
    Dim Transaction_ID As Double
    Transaction_ID = val(Me.Text2.text)

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
        
    If Me.DcAccount1.BoundText = "" Then
        Me.DcAccount1.BoundText = Account_Code_dynamic
    End If

    If Me.DcAccount2.BoundText = "" Then
        Me.DcAccount2.BoundText = Account_Code_dynamic
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
Private Sub TxtShortName_KeyDown(KeyCode As Integer, Shift As Integer)
'   LoadSpecificItems
SerchItems (TxtShortName.text)
DoEvents
DoEvents
DoEvents
DoEvents

        If KeyCode = vbKeyReturn Then
        
        
   DCboItemsName.SetFocus
   DCboItemsName.BoundText = ""
        Sendkeys "{F4}"
        End If
End Sub
Sub SerchItems(Optional str As String)
 
Set DCboItemsCode.RowSource = Nothing
Set DCboItemsName.RowSource = Nothing
If str <> "" Then
Dim sql As String
Dim SQL1 As String
 
Dim StrWhere As String
  Dim astrSplit2tems2() As String
  Dim j As Integer
  Dim nElements As Integer
  Dim SearchString As String
StrWhere = ""
SearchString = ""
sql = " select  ItemID,barCodeNO   from  dbo.TblItems where 1=1"
If SystemOptions.UserInterface = ArabicInterface Then
SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where 1=1"
Else
SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where 1=1"
End If

          astrSplit2tems2 = Split(str, " ")
          nElements = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
          If nElements = 0 Then
                         If SystemOptions.UserInterface = ArabicInterface Then
                            StrWhere = " and (ItemName Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%'  or fullcode Like N'%" & Trim(str) & "%') "
                    Else
                            StrWhere = " and (ItemNamee Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%' or fullcode Like N'%" & Trim(str) & "%' ) "
                    End If
                    
          End If
        If nElements > 0 Then
        
     '   StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(0)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(0)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(0)) & "%') "
        SearchString = ""
        For j = 0 To nElements
        
         SearchString = SearchString & "%" & Trim(astrSplit2tems2(j))
             '     SearchString = "%" & Trim(astrSplit2tems2(j)) & SearchString
                  
        '   StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(j)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(j)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(j)) & "%') "
        '   StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
         Next j
         SearchString = SearchString & "%"
                             If SystemOptions.UserInterface = ArabicInterface Then

             StrWhere = StrWhere + " and (ItemName Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
             Else
              StrWhere = StrWhere + " and (ItemNamee Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
             End If
        '-  StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
      
         End If
        
    sql = sql & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        sql = sql + " Order BY ItemName "
    Else
        sql = sql + " Order BY ItemName "
    End If


    SQL1 = SQL1 & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        SQL1 = SQL1 + " Order BY ItemName "
    Else
        SQL1 = SQL1 + " Order BY ItemNamee "
    End If
    
   End If
    fill_combo DCboItemsCode, sql
        fill_combo DCboItemsName, SQL1
        DoEvents
        DoEvents
  
                        If str = "" Then
                                 sql = " select  ItemID,barCodeNO   from  dbo.TblItems where 1=1"
                                 If SystemOptions.UserInterface = ArabicInterface Then
                                 SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where 1=1"
                                     SQL1 = SQL1 + " Order BY ItemName "
                                 Else
                                 SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where 1=1"
                                     SQL1 = SQL1 + " Order BY ItemNameE "
                                 End If
                                 
                                     fill_combo DCboItemsCode, sql
                                         fill_combo DCboItemsName, SQL1
                End If
                
       Exit Sub
       
If str <> "" Then
'Dim Sql As String
'Dim StrWhere As String
'  Dim astrSplit2tems2() As String
'  Dim j As Integer
'  Dim nElements As Integer
StrWhere = ""
If SystemOptions.UserInterface = ArabicInterface Then
sql = " select  ItemID,ItemName   from  dbo.TblItems where 1=1"
Else
sql = " select  ItemID,ItemNamee   from  dbo.TblItems where 1=1"
End If
          astrSplit2tems2 = Split(str, " ")
          nElements = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
        If nElements > 0 Then
        StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(0)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(0)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(0)) & "%') "
        For j = 1 To nElements - 1
        
           StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(j)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(j)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(j)) & "%') "
           StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
         Next j
         End If
    sql = sql & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        sql = sql + " Order BY ItemName "
    Else
        sql = sql + " Order BY ItemNamee "
    End If


   End If
   
        fill_combo DCboItemsName, sql
        
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
        If Row = .rows - 1 Then
            .rows = .rows + 1
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

If SystemOptions.AllowSett = True Then
CMDStartGard.Enabled = True
End If

If SystemOptions.AllowSett1 = True Then
CMDStartSetelment.Enabled = True
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
    Set NewGrid.VatGrid = Me.VatGrid
    Set NewGrid.TxtValueAdded = TxtValueAdded
    
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
        Dcombos.GetAccountingCodes DcAccount1, True
        Dcombos.GetAccountingCodes DcAccount2, True
    Else
 
        Dcombos.GetAccountingCodesENg DcAccount1, True
        Dcombos.GetAccountingCodesENg DcAccount2, True

    End If

    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboStoreName

    StrSQL = "Select * From Transactions where Transaction_Type=30"
    StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
    StrSQL = StrSQL & "   Order By Transaction_Date ,StoreID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    XPBtnMove_Click 2
    chkIsAutoCost.Enabled = True
    
    TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
    Msg = Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRight, App.Title
End Sub

Private Sub WriteDev()
    On Error GoTo errortrap
    Dim Account_Code_dynamic As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

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

Public Function retrive1(Optional StoreID As Integer, _
                         Optional FromDate As Date, _
                         Optional ToDate As Date)
    Dim StrSQL As String
    Dim RsDetails As New ADODB.Recordset
    Dim RowNum As Integer
    Dim LngNoteID As Long

    On Error GoTo ErrTrap
 
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    StrSQL = "SELECT     TOP 100 PERCENT SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.Transaction_Details.Item_ID AS ItemID, "
    StrSQL = StrSQL & "  dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.Transaction_Details.UnitId, dbo.TblItems.ItemCode,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemName, dbo.TblUnites.UnitName, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ClassId,"
    StrSQL = StrSQL & "  dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblItemsSizes.SizeName AS SizeName, dbo.TblItemsColors.ColorName ,dbo.Transaction_Details.lotNo"
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
    StrSQL = StrSQL & "   AND (dbo.Transactions.StoreID =" & StoreID & ")"
    StrSQL = StrSQL & "  GROUP BY dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID, dbo.TblStore.StoreName, dbo.Transaction_Details.order_no, dbo.Transaction_Details.UnitId,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblUnites.UnitName, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize,"
    StrSQL = StrSQL & "   dbo.Transaction_Details.ClassId , dbo.TblItemsclasses.SizeName, dbo.TblItemsSizes.SizeName, dbo.TblItemsColors.ColorName,dbo.Transaction_Details.lotNo"
    StrSQL = StrSQL & "  Having (SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) <> 0)"
    StrSQL = StrSQL & "  ORDER BY dbo.TblItems.ItemID"


    Dim LngItemID As Long
    Dim LngUnitID As Long
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

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
            FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
            FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), 1, (RsDetails("UnitID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
      
            RsDetails.MoveNext

            If FG.rows > 10 Then
                If RowNum = 8 Then FG.Refresh
            End If

        Next RowNum

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    Me.XPTxtSum.text = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Valu"), FG.rows - 1, FG.ColIndex("Valu"))
    Me.LblTotalQty = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Count"), FG.rows - 1, FG.ColIndex("Count"))

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

Private Sub Cmd_Click(Index As Integer)
     On Error GoTo ErrTrap
    Dim intDef As Integer

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            Me.TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=30"))
            txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
            WriteDev
            GridDefaultValue FG.rows - 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
            FG.SetFocus
            FG.rows = 2
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.rows - 1
            Me.dcBranch.BoundText = branch_id
            Dim FirstPeriodDateInthisYear  As Date
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            Me.DTPickerAccFrom = FirstPeriodDateInthisYear
            DTPickerAccTo.value = Date
            opt(2).value = True

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            'If AvailableDeal = True Then
            TxtModFlg.text = "E"
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

    Select Case Me.TxtModFlg.text

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
            
        'Wael
        '    Fg.Editable = flexEDNone

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

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBillID.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
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
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«   ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· Â–Â «·»Ì«‰« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ Â–Â «·»Ì«‰« " & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› Â–Â «·»Ì«‰« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
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

    Select Case TxtModFlg.text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                XPBtnMove_Click (1)
            
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                rs.Find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Exit Sub
                End If

                Me.TxtModFlg.text = "R"
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
    
    If Me.TxtModFlg.text <> "R" Then
        If DCboStoreName.BoundText = "" Then
            Msg = "ÌÃ» «Œ Ì«— «”„ «·„Œ“‰"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboStoreName.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If NewGrid.IsReaptedSerials = True Then
            Msg = "ÌÊÃœ  ﬂ—«— ›Ï √—ﬁ«„ «·”Ì—Ì«· «·„œŒ·… "
            Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ «·√—ﬁ«„ «·„œŒ·…"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        '-----------------------------------------------
        'Check the Items Grid
        If chkIsAutoCost.value = vbUnchecked Then
        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If
        End If

        '----------------------------------------------
        Cn.BeginTrans
        BegineTrans = True

        If TxtModFlg.text = "N" Then
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            Me.TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=30"))
            rs.AddNew
            rs("Transaction_ID").value = val(XPTxtBillID.text)
        End If

    '    RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
        rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
        rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)
        rs("Transaction_Serial").value = Me.TxtTransSerial.text
        rs("Transaction_Date").value = XPDtbBill.value
    
        rs("GardFromDate").value = DTPickerAccFrom.value
        rs("GardTodate").value = DTPickerAccTo.value

        If opt(0).value = True Then
            rs("GardEntryType").value = 0
        ElseIf opt(1).value = True Then
            rs("GardEntryType").value = 1
        ElseIf opt(2).value = True Then
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
        rs("Account1").value = IIf(DcAccount1.BoundText = "", Null, DcAccount1.BoundText)
        rs("Account2").value = IIf(DcAccount2.BoundText = "", Null, DcAccount2.BoundText)
    If chkAutoDetect.value = vbChecked Then
    rs("chkAutoDetect").value = 1
    Else
    rs("chkAutoDetect").value = 0
    
    End If
    
        rs.update

        If Me.TxtModFlg.text = "E" Then
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            ' StrSqlDel = "delete From NOTES where Transaction_ID=" & Val(rs("Transaction_ID").value)
            ' Cn.Execute StrSqlDel, , adExecuteNoRecords
    
            StrSqlDel = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
    
        End If

        For RowNum = 1 To FG.rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
RSTransDetails("AutoDetect").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("AutoDetect")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("AutoDetect"))))

                RSTransDetails("Transaction_ID").value = XPTxtBillID.text
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
                         RSTransDetails("ProductionDate").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")))
                RSTransDetails("ExpiryDate").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")))
               
                RSTransDetails("GardQty").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("GardQty")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("GardQty"))))
                RSTransDetails("Gardresult").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult"))))
                RSTransDetails("Gardresult1").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult1")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult1"))))
                RSTransDetails("Gardresult2").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult2")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Gardresult2"))))
                RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))

                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
    RSTransDetails("ParrtNoCode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))))
      RSTransDetails("ItemDetailedCode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))))
      
                RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
                ' IIf((FG.TextMatrix(RowNum, FG.ColIndex("BranchId")) = ""), 1, Val(FG.TextMatrix(RowNum, FG.ColIndex("BranchId"))))
               
                ' RSTransDetails("ItemSize").value = _
                  IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), "", Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
                RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))

                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
        
                LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                     RSTransDetails("QtyBySmalltUnit").value = IIf(IsNull(RsUnitData("UnitFactor").value), 1, RsUnitData("UnitFactor").value)
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
        Me.lblAccountInterval.Caption = SystemOptions.SysCurrentAccountIntervalID
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
      
        Select Case Me.TxtModFlg.text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Successfully Saved " & CHR(13)
                    Msg = Msg + "Do you want to enter another  New operation"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox "Successfully Updated", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
            
        End Select

        TxtModFlg.text = "R"
 
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  "
    Msg = Msg & CHR(13) & "" & Err.Description
    Msg = Msg & CHR(13) & "" & Err.Number
    Msg = Msg & CHR(13) & "" & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Function Coloring()
Dim LngRow As Long
With Me.FG
For LngRow = 1 To FG.rows - 1

        If val(FG.TextMatrix(LngRow, FG.ColIndex("AutoDetect"))) = 0 Then
   
    .cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = 0
   Else
    .cell(flexcpBackColor, LngRow, 1, LngRow, .Cols - 1) = vbRed
     End If
  Next LngRow
     
 End With
NewGrid.CountItems
            
End Function

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim StrSQL As String
    Dim RsDetails As New ADODB.Recordset
    Dim RowNum As Integer
    Dim LngNoteID As Long

 '   On Error GoTo ErrTrap

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.Find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
'Me.TxtModFlg.text = "R"

    Screen.MousePointer = vbArrowHourglass
    opt2(0).value = True
    Opt3(0).value = True
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", val(rs("BranchId").value))
    Text1.text = IIf(IsNull(rs("NotS").value), "", (rs("NotS").value))
    Text2.text = IIf(IsNull(rs("NotS2").value), "", (rs("NotS2").value))

    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
    txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)
    Me.TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", rs("Transaction_Date").value)

    DTPickerAccFrom.value = IIf(IsNull(rs("GardFromDate").value), "", rs("GardFromDate").value)
    DTPickerAccTo.value = IIf(IsNull(rs("GardTodate").value), "", rs("GardTodate").value)

    If IsNull(rs("GardEntryType").value) Then
        opt(2).value = True
    Else
        opt(val(rs("GardEntryType").value)).value = True

    End If

    If IsNull(rs("StartGard").value) Then
        ChkStartGard.value = vbUnchecked
    ElseIf (rs("StartGard").value) = True Then
        ChkStartGard.value = vbChecked
    ElseIf (rs("StartGard").value) = False Then
        ChkStartGard.value = vbUnchecked

    End If



    If IsNull(rs("chkAutoDetect").value) Then
        chkAutoDetect.value = vbUnchecked
    ElseIf (rs("chkAutoDetect").value) = True Then
        chkAutoDetect.value = vbChecked
    ElseIf (rs("chkAutoDetect").value) = False Then
        chkAutoDetect.value = vbUnchecked

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

    DcAccount1.BoundText = IIf(IsNull(rs("Account1").value), "", rs("Account1").value)
    DcAccount2.BoundText = IIf(IsNull(rs("Account2").value), "", rs("Account2").value)
    DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
'    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
'    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
 
    
        StrSQL = "SELECT  ProductionDate,ExpiryDate, AutoDetect,   dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.TblItems.HaveSerial, dbo.TblItems.LastUpdate, "
  StrSQL = StrSQL + "                     dbo.TblItems.UserID, dbo.TblItems.PurchasePrice, dbo.TblItems.SallingPrice, dbo.TblItems.RequestLimit, dbo.TblItems.CustomerPrice, dbo.TblItems.DealerPrice,"
 StrSQL = StrSQL + "                      dbo.TblItems.HaveGuarantee, dbo.TblItems.GuaranteeValue, dbo.TblItems.GuaranteeType, dbo.TblItems.IsArchive, dbo.TblItems.ItemType, dbo.TblItems.AssbliedItem,"
 StrSQL = StrSQL + "                      dbo.TblItems.RelatedItem, dbo.TblItems.ItemComment, dbo.TblItems.ItemCase, dbo.TblItems.ItemMaking, dbo.TblItems.ItemMakingNew, dbo.TblItems.code,"
 StrSQL = StrSQL + "                      dbo.TblItems.Branch_NO, dbo.TblItems.Fullcode, dbo.TblItems.prifix, dbo.TblItems.PartNo, dbo.TblItems.CostPrice, dbo.TblItems.ItemNamee,"
 StrSQL = StrSQL + "                      dbo.TblItems.DefaultSupplier, dbo.TblItems.itemSerials, dbo.TblItems.BinLocation, dbo.TblItems.minvalueqty, dbo.TblItems.MaxValueqty, dbo.TblItems.FreeQty,"
 StrSQL = StrSQL + "                      dbo.TblItems.barCodeNO, dbo.TblItems.CatlogNO, dbo.TblItems.FactoryNO, dbo.TblItems.TemplateID, dbo.TblItems.ItemMaxDiscount, dbo.TblItems.OverHead,"
 StrSQL = StrSQL + "                      dbo.TblItems.Wight, dbo.TblItems.Content, dbo.TblItems.Dippre, dbo.TblItems.Source, dbo.TblItems.Typenew, dbo.TblItems.maxRecivePeriod,"
 StrSQL = StrSQL + "                      dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transaction_Details.ItemCase AS Expr1,"
 StrSQL = StrSQL + "                      dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount,"
 StrSQL = StrSQL + "                      dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.Remarks, dbo.Transaction_Details.ShowQty,"
 StrSQL = StrSQL + "                      dbo.Transaction_Details.UnitId, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ClassId, dbo.Transaction_Details.ProductionDate,"
 StrSQL = StrSQL + "                      dbo.Transaction_Details.ItemDetailedCode ,ParrtNoCode,GardQty,Gardresult,Gardresult1,Gardresult2,dbo.Transaction_Details.lotNo"
 StrSQL = StrSQL + " FROM         dbo.TblItems INNER JOIN"
 StrSQL = StrSQL + "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID LEFT OUTER JOIN"
 StrSQL = StrSQL + "                      dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"

'    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
 
    
        
    StrSQL = StrSQL + " order by Transaction_Details.id "

    
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Dim ItemID As Long
    Dim UnitID As Long
    Dim itemsize As Long
    Dim ColorID As Long
    Dim ClassId As Long
 
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For RowNum = 1 To RsDetails.RecordCount

            With FG
                ItemID = IIf(IsNull(RsDetails("Item_ID").value), 0, RsDetails("Item_ID").value)
        
                UnitID = IIf(IsNull(RsDetails("UnitID")), 1, (RsDetails("UnitID").value))
                itemsize = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
                ColorID = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
                ClassId = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
               FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
        .TextMatrix(RowNum, FG.ColIndex("AutoDetect")) = IIf(IsNull(RsDetails("AutoDetect").value), 0, RsDetails("AutoDetect").value)
                .TextMatrix(RowNum, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID").value), "", RsDetails("Item_ID").value)
                .TextMatrix(RowNum, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID").value), "", RsDetails("Item_ID").value)
                .TextMatrix(RowNum, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Showqty").value), "", RsDetails("Showqty").value)
             'GardQty,Gardresult,Gardresult1,Gardresult2
                .TextMatrix(RowNum, FG.ColIndex("GardQty")) = IIf(IsNull(RsDetails("GardQty").value), "", RsDetails("GardQty").value)
                .TextMatrix(RowNum, FG.ColIndex("Gardresult")) = IIf(IsNull(RsDetails("Gardresult").value), "", RsDetails("Gardresult").value)
                .TextMatrix(RowNum, FG.ColIndex("Gardresult1")) = IIf(IsNull(RsDetails("Gardresult1").value), "", RsDetails("Gardresult1").value)
                .TextMatrix(RowNum, FG.ColIndex("Gardresult2")) = IIf(IsNull(RsDetails("Gardresult2").value), "", RsDetails("Gardresult2").value)
             .TextMatrix(RowNum, .ColIndex("ParrtNoCode")) = IIf(IsNull(RsDetails("ParrtNoCode")), "", (RsDetails("ParrtNoCode").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode")) = IIf(IsNull(RsDetails("ItemDetailedCode")), "", (RsDetails("ItemDetailedCode").value))
            
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
            '
              FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            

            FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
      
            'FG.TextMatrix(RowNum, FG.ColIndex("GardQty")) = GetActualItemQty(Val(Me.DCboStoreName.BoundText), Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value, _
             ItemID, UnitId, ItemSize, ColorID, ClassId)
            'FG.TextMatrix(RowNum, FG.ColIndex("Gardresult")) = Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) - Val(FG.TextMatrix(RowNum, FG.ColIndex("GardQty")))
            ' «·›—ﬁ Ì”«ÊÌ „«  „ «œŒ«·Â
            '„ÿ—ÊÕ „‰Â «·„ÊÃÊœ ›⁄·« »«·»—‰«„Ã
 
            RsDetails.MoveNext

            If FG.rows > 10 Then
                If RowNum = 8 Then FG.Refresh
            End If

        Next RowNum

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    Me.XPTxtSum.text = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Valu"), FG.rows - 1, FG.ColIndex("Valu"))
    Me.LblTotalQty = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Count"), FG.rows - 1, FG.ColIndex("Count"))
 
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    NewGrid.CountItems
    Coloring
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub printing()
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Set BalanceReport = New ClsOpeningBalanceReport
        BalanceReport.ShowOpeningBalanceData XPTxtBillID.text
    End If

    Exit Sub
ErrTrap:
End Sub
 Public Function GardReport(Transactionid As Double, _
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
    StrSQL = StrSQL & "  dbo.TblStore.StoreName , dbo.TblItems.SallingPrice , dbo.Transaction_Details.LotNO"
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
      ElseIf ReportType = 3 Then
        StrSQL = StrSQL & " AND (dbo.Transaction_Details.Gardresult = 0)"
  
  
    End If
    

'    StrSQL = StrSQL & "  ORDER BY dbo.Transaction_Details.Item_ID"
StrSQL = StrSQL + "order by dbo.Transaction_Details.id"

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Dir(App.path & "\Reports\Inventory\Gard3.rpt") = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg + Reportpath
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(5).AddCurrentValue GetCurrentGardEmployee(val(DCboStoreName.BoundText))
          
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
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            StrSQL = "select * From QryDelPurchase where Transaction_Date>=" & SQLDate(XPDtbBill.value, True) & ""
            StrSQL = StrSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))

            '        If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) <> "" Then
            '            If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = True Then
            If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                End If

                '            End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then

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
                    Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, val(Me.XPTxtBillID.text))

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
        If Me.TxtModFlg.text = "R" Then
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
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
        
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
        
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            
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
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Stock Settlement Auto"
    C1Elastic6.Caption = Me.Caption
    Ele(2).Caption = "Period"
    lbl(10).Caption = "From"
    lbl(11).Caption = "To"
    lbl(1).Caption = "ID"
    lbl(0).Caption = "Date"
    Label3.Caption = "Branch"

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
   
    On Error Resume Next

    With Me.FG
        .TextMatrix(0, .ColIndex("Gardresult")) = "Diff"
        .TextMatrix(0, .ColIndex("Gardresult1")) = "Dec"
        .TextMatrix(0, .ColIndex("Gardresult2")) = "Inc"
        .TextMatrix(0, .ColIndex("GardQty")) = "Current Qty"
    End With

    CMDStartSetelment.Caption = "Start Settlement"

    'NewItem

End Sub

Private Sub XPTxtSum_Change()
    Me.LblTotal.Caption = XPTxtSum.text
    Exit Sub
ErrTrap:
End Sub
