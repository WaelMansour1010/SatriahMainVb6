VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmNewGard 
   Caption         =   "«œŒ«· «·Ã—œ «·›⁄·Ï ··„Œ«“‰"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13395
   HelpContextID   =   90
   Icon            =   "FrmNewGard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmNewGard.frx":038A
   RightToLeft     =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   13395
   WindowState     =   2  'Maximized
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
      Height          =   7845
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13395
      _cx             =   23627
      _cy             =   13838
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
      GridRows        =   6
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmNewGard.frx":0714
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4785
         Index           =   3
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2040
         Width           =   13365
         _cx             =   23574
         _cy             =   8440
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
         GridRows        =   6
         GridCols        =   6
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmNewGard.frx":078C
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin MSComctlLib.Toolbar TBr 
            Height          =   630
            Left            =   510
            TabIndex        =   12
            Top             =   4155
            Width           =   12330
            _ExtentX        =   21749
            _ExtentY        =   1111
            ButtonWidth     =   609
            ButtonHeight    =   1005
            Appearance      =   1
            _Version        =   393216
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   885
            Index           =   4
            Left            =   30
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   30
            Width           =   13305
            _cx             =   23469
            _cy             =   1561
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
               Height          =   270
               Left            =   600
               TabIndex        =   96
               Top             =   120
               Width           =   6795
            End
            Begin VB.TextBox TxtItemsIDes 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   11160
               TabIndex        =   89
               Top             =   0
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.TextBox TxtPrice 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   675
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   630
               Width           =   1635
            End
            Begin VB.TextBox TxtSerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   285
               Left            =   4050
               MaxLength       =   20
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   630
               Width           =   2055
            End
            Begin VB.TextBox TxtQuantity 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   2400
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   630
               Width           =   1530
            End
            Begin VB.ComboBox CboItemCase 
               Height          =   315
               Left            =   6150
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   630
               Width           =   1905
            End
            Begin MSDataListLib.DataCombo DCboItemsName 
               Height          =   315
               Left            =   8055
               TabIndex        =   37
               Top             =   630
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboItemsCode 
               Height          =   315
               Left            =   10650
               TabIndex        =   36
               Top             =   630
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdAdd 
               Height          =   405
               Left            =   30
               TabIndex        =   47
               Top             =   510
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   714
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
               ButtonImage     =   "FrmNewGard.frx":0829
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
               Height          =   285
               Index           =   97
               Left            =   7830
               TabIndex        =   97
               Top             =   120
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”⁄—"
               Height          =   255
               Index           =   26
               Left            =   855
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   375
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   255
               Index           =   27
               Left            =   2655
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   375
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ì—Ì«·"
               Height          =   360
               Index           =   28
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   375
               Width           =   1950
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·’‰›"
               Height          =   255
               Index           =   29
               Left            =   6285
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   375
               Width           =   1770
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈”„ «·’‰›"
               Height          =   255
               Index           =   30
               Left            =   8280
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   375
               Width           =   2370
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   255
               Index           =   31
               Left            =   10875
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   375
               Width           =   2415
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   3105
            Left            =   30
            TabIndex        =   35
            Top             =   1035
            Width           =   13305
            _cx             =   23469
            _cy             =   5477
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
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmNewGard.frx":0BC3
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
            Height          =   465
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   4155
            Width           =   450
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1380
         Index           =   1
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   645
         Width           =   13335
         _cx             =   23521
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
         Begin VB.CheckBox chkIsCostFromExcel 
            Alignment       =   1  'Right Justify
            Caption         =   "«· ﬂ·›… „‰ «·Ã—œ"
            Height          =   315
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   690
            Width           =   1395
         End
         Begin VB.CheckBox chkBigCost 
            Caption         =   "«’‰«› » ﬂ«·Ì› ﬂ»Ì—…"
            Height          =   255
            Left            =   8370
            TabIndex        =   98
            Top             =   1050
            Width           =   1935
         End
         Begin VB.TextBox txtStoreSearch 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10800
            TabIndex        =   94
            Top             =   690
            Width           =   1125
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Õ–› «·«’‰«› «·«·Ì…"
            Height          =   375
            Left            =   6900
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   1020
            Width           =   1455
         End
         Begin VB.TextBox TxtItemCodeB 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10515
            TabIndex        =   60
            Top             =   1080
            Width           =   1440
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   480
            Width           =   1200
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   1200
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Frame Frame1 
            Caption         =   "Õœœ ÿ—Ìﬁ… «·«œŒ«·"
            Height          =   1335
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   -90
            Width           =   3855
            Begin VB.OptionButton OPT 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬂ· «’‰«› «·„Œ“‰ »«·’·«ÕÌ…"
               Height          =   195
               Index           =   4
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   480
               Width           =   2295
            End
            Begin VB.CheckBox chkNew 
               Alignment       =   1  'Right Justify
               Caption         =   " — Ì» ÃœÌœ"
               Height          =   195
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   1080
               Width           =   1875
            End
            Begin VB.OptionButton OPT 
               Alignment       =   1  'Right Justify
               Caption         =   "„Ã„Ê⁄Â „Õœœ…"
               Height          =   195
               Index           =   3
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox txtFile 
               Height          =   285
               Left            =   720
               Locked          =   -1  'True
               TabIndex        =   64
               Top             =   -135
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CommandButton Command2 
               Caption         =   " Õ„Ì· «·„·›..."
               Height          =   255
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton Command1 
               Caption         =   " ÕœÌœ «·„·›..."
               Height          =   255
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton OPT 
               Alignment       =   1  'Right Justify
               Caption         =   "ÌœÊÌ"
               Height          =   195
               Index           =   2
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   960
               Width           =   1215
            End
            Begin VB.OptionButton OPT 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬂ· «’‰«› «·„Œ“‰"
               Height          =   195
               Index           =   1
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   480
               Width           =   1575
            End
            Begin VB.OptionButton OPT 
               Alignment       =   1  'Right Justify
               Caption         =   "„‰ „·›"
               Height          =   195
               Index           =   0
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   240
               Width           =   1575
            End
            Begin MSComDlg.CommonDialog CD1 
               Left            =   0
               Top             =   1200
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin MSDataListLib.DataCombo XPCboGroup 
               Height          =   315
               Left            =   120
               TabIndex        =   66
               Top             =   720
               Width           =   2160
               _ExtentX        =   3810
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.TextBox txtopening_balance_voucher_id 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2910
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   1560
            Visible         =   0   'False
            Width           =   1110
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
            Left            =   -7830
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   0
            Visible         =   0   'False
            Width           =   7920
            Begin MSDataListLib.DataCombo DcboDebitSide 
               Height          =   315
               Left            =   1680
               TabIndex        =   19
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
               TabIndex        =   20
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
               TabIndex        =   26
               Top             =   510
               Width           =   1095
            End
            Begin VB.Label LblDevID 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   5580
               RightToLeft     =   -1  'True
               TabIndex        =   25
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
               TabIndex        =   24
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
               TabIndex        =   23
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
               TabIndex        =   22
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
               TabIndex        =   21
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
            TabIndex        =   15
            Top             =   1170
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   60
            Width           =   600
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   1020
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   630
            Visible         =   0   'False
            Width           =   825
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   345
            Left            =   9360
            TabIndex        =   58
            Top             =   120
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   609
            _Version        =   393216
            Format          =   237109251
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   8235
            TabIndex        =   57
            Top             =   690
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   705
            Index           =   2
            Left            =   6840
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   0
            Width           =   2475
            _cx             =   4366
            _cy             =   1244
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
            ForeColor       =   16711680
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   " ÕœÌœ «·› —… «·“„‰Ì…"
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
               Left            =   2850
               TabIndex        =   56
               ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
               Top             =   720
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   -2147483624
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   237109251
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
               TabIndex        =   59
               ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
               Top             =   240
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   -2147483624
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   237109251
               CurrentDate     =   37357
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·Ã—œ"
               ForeColor       =   &H00FF8080&
               Height          =   285
               Index           =   11
               Left            =   1590
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   240
               Width           =   795
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   285
               Index           =   10
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   765
               Visible         =   0   'False
               Width           =   555
            End
         End
         Begin ImpulseButton.ISButton SearchCashCustomer 
            Height          =   315
            Index           =   1
            Left            =   10275
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1080
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
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
            ButtonImage     =   "FrmNewGard.frx":0E66
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»«—ﬂÊœ"
            Height          =   375
            Index           =   12
            Left            =   12270
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·’›Õ… »„Õ÷— «·Ã—œ"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1320
            TabIndex        =   54
            Top             =   480
            Width           =   1650
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   375
            Index           =   2
            Left            =   12285
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2280
            TabIndex        =   34
            Top             =   120
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«œŒ«·"
            Height          =   375
            Index           =   0
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   105
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”·”·"
            Height          =   375
            Index           =   1
            Left            =   11520
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   75
            Width           =   1680
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   615
         Left            =   15
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   15
         Width           =   13365
         _cx             =   23574
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
         Caption         =   "«œŒ«· «·Ã—œ «·›⁄·Ï ··„Œ«“‰ "
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
            Left            =   8040
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Text            =   "0"
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1725
            TabIndex        =   28
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
            ButtonImage     =   "FrmNewGard.frx":1263
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
            TabIndex        =   29
            Top             =   120
            Width           =   750
            _ExtentX        =   1323
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
            ButtonImage     =   "FrmNewGard.frx":15FD
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
            Left            =   2670
            TabIndex        =   30
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
            ButtonImage     =   "FrmNewGard.frx":1997
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
            TabIndex        =   31
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
            ButtonImage     =   "FrmNewGard.frx":1D31
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
            Left            =   3720
            TabIndex        =   90
            Tag             =   "1"
            Top             =   0
            Visible         =   0   'False
            Width           =   4080
            _cx             =   7197
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
            FormatString    =   $"FrmNewGard.frx":20CB
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   5
         Left            =   15
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   6840
         Width           =   13335
         _cx             =   23521
         _cy             =   767
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
            Height          =   285
            Left            =   4365
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   105
            Visible         =   0   'False
            Width           =   60
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   2100
            TabIndex        =   69
            Top             =   105
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   315
            Index           =   63
            Left            =   4665
            TabIndex        =   78
            Top             =   120
            Width           =   510
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
            Height          =   330
            Left            =   4125
            TabIndex        =   77
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5745
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   105
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·—’Ìœ"
            Height          =   210
            Index           =   3
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   120
            Width           =   465
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   285
            Index           =   6
            Left            =   3540
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   105
            Width           =   540
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   210
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   120
            Width           =   300
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   225
            Left            =   1155
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   105
            Width           =   120
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   330
            Index           =   5
            Left            =   495
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   135
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   345
            Index           =   4
            Left            =   1275
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   135
            Width           =   750
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   0
         Left            =   0
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   7305
         Width           =   13395
         _cx             =   23627
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
            Height          =   360
            Index           =   0
            Left            =   12105
            TabIndex        =   80
            Top             =   105
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
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
            Height          =   360
            Index           =   1
            Left            =   10605
            TabIndex        =   81
            Top             =   105
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   635
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
            Height          =   360
            Index           =   2
            Left            =   9015
            TabIndex        =   82
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
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
            Height          =   360
            Index           =   3
            Left            =   7515
            TabIndex        =   83
            Top             =   105
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
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
            Height          =   360
            Index           =   4
            Left            =   5745
            TabIndex        =   84
            Top             =   105
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
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
            Height          =   360
            Index           =   5
            Left            =   4455
            TabIndex        =   85
            Top             =   105
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   635
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
            Height          =   360
            Index           =   6
            Left            =   30
            TabIndex        =   86
            Top             =   105
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   635
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
            Height          =   360
            Index           =   7
            Left            =   2955
            TabIndex        =   87
            Top             =   105
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   635
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
            Height          =   360
            Left            =   1635
            TabIndex        =   88
            Top             =   105
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   635
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
Attribute VB_Name = "FrmNewGard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim BalanceReport As ClsOpeningBalanceReport
Dim cSearchDcbo As clsDCboSearch
Dim NewGrid As New ClsGrid
Dim error_string As String
Dim RsTemp As ADODB.Recordset



Private Sub chkBigCost_Click()
retrive1 val(DCboStoreName.BoundText), DTPickerAccFrom.value, DTPickerAccTo.value, 10
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub Command1_Click()
        If DCboStoreName.BoundText = "" And chkNew.value = vbUnchecked Then
            Msg = "ÌÃ» «Œ Ì«— «”„ «·„Œ“‰"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboStoreName.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
CD1.ShowOpen
txtFile.text = CD1.FileName
End Sub
Sub FillItemDete()
'Â–« «·ﬂÊœ ›Ì Õ«·Â «· ⁄«„· » ›«’Ì· «·«’‰«›
  error_string = ""

   

StrSQL = "SELECT     dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ProductionDate, dbo.ItemsDetails.ExpireDate, dbo.ItemsDetails.ColorID, "
StrSQL = StrSQL & "                       dbo.ItemsDetails.unitid , dbo.ItemsDetails.sizeid, dbo.ItemsDetails.ClassId, dbo.ItemsDetails.ItemID, dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName"
StrSQL = StrSQL & " FROM         dbo.ItemsDetails LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID"
StrSQL = StrSQL & "  GROUP BY dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.UnitID, dbo.ItemsDetails.SizeID,"
StrSQL = StrSQL & "                        dbo.ItemsDetails.ClassId , dbo.ItemsDetails.ProductionDate, dbo.ItemsDetails.ExpireDate, dbo.ItemsDetails.ItemID, dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName"


'StrSQL = StrSQL & "  HAVING        (ParrtNoCode = '" & Barcode & "')"
     
     
       
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText




If txtFile.text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
Dim currentvalue As String

Dim itemcode As String
Dim itemqty As Double
Dim des As String
Dim DebitValue As String
Dim CreditValue As String
  

    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")

    ExcelObj.Workbooks.Open txtFile.text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 
    With ExcelSheet
    i = 2
    Do Until .cells(i, 2) & "" = ""
 '       Set l = lvwList.ListItems.Add(, , .Cells(i, 1))
    itemcode = .cells(i, 1)
    itemqty = .cells(i, 2)
     
        
 addrow itemcode, itemqty
       i = i + 1
       NewGrid.CountItems
    Loop
        End With
    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing

        If error_string <> "" Then
            CreatLog_File_for_error (error_string)
       End If
'NewGrid.Calculate 1, , , True
GetNotinGard
Coloring
End Sub
Sub FillItem()
  error_string = ""
If txtFile.text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
    Dim currentvalue As String
    Dim Name As String
    Dim itemcode As String
    Dim itemqty As Double
    Dim des As String
    Dim DebitValue As String
    Dim CreditValue As String
  
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")
    ExcelObj.Workbooks.Open txtFile.text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 
    With ExcelSheet
    i = 2
    
    Do Until .cells(i, 1) & "" = ""
    itemcode = .cells(i, 1)
    itemqty = .cells(i, 3)
    Name = .cells(i, 4)
        
 addrow2 itemcode, itemqty, Name
       i = i + 1
       NewGrid.CountItems
    Loop
        End With
    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing

        If error_string <> "" Then
            CreatLog_File_for_error (error_string)
       End If
GetNotinGard
Coloring
End Sub

Sub FillItemNew()
    Dim excelApp As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Long

    Dim itemcode As String
    Dim Name As String
    Dim itemqty As Double
    Dim UnitName As String
    Dim ExpireDate As String
    Dim mPrice As Double
    Dim sizename As String
    Dim colorname As String
    Dim mWidth As String
    Dim mLength As String
    Dim mHeight As String
    Dim mArea As String

    On Error GoTo errHandler
    error_string = ""

    If txtFile.text = "" Then
        MsgBox "Õœœ «·„·› √Ê·«"
        Exit Sub
    End If

    ' «› Õ Excel „—… Ê«Õœ…
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False

    Set ExcelBook = excelApp.Workbooks.Open(txtFile.text)
    Set ExcelSheet = ExcelBook.Worksheets(1)

    i = 2
    With ExcelSheet
        Do While (.cells(i, 1).value & "" <> "") Or (.cells(i, 2).value & "" <> "")
            itemcode = .cells(i, 1).value & ""
            Name = .cells(i, 2).value & ""
            itemqty = val(.cells(i, 3).value)
            UnitName = .cells(i, 4).value & ""
            ExpireDate = .cells(i, 5).value & ""

            sizename = Trim(.cells(i, 6).value & "")
            colorname = Trim(.cells(i, 7).value & "")
            mWidth = Trim(.cells(i, 8).value & "")
            mLength = Trim(.cells(i, 9).value & "")
            mHeight = Trim(.cells(i, 10).value & "")
            mArea = Trim(.cells(i, 11).value & "")

            ' ·Ê ⁄‰œﬂ ”⁄— ›Ì «·⁄„Êœ «·—«»⁄ √Ê €Ì—Â ÕÿÂ Â‰«
            ' mPrice = Val(.Cells(i, 12).Value)
mPrice = 0
            addrow2 itemcode, itemqty, UnitName, mPrice, ExpireDate, _
                    sizename, colorname, mWidth, mLength, mHeight, mArea

            NewGrid.CountItems

            i = i + 1
            DoEvents          ' „Â„ ⁄‘«‰ „« Ìÿ·⁄‘ Component Busy
        Loop
    End With

CleanExit:
    On Error Resume Next
    If Not ExcelBook Is Nothing Then ExcelBook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set excelApp = Nothing

    If error_string <> "" Then
        CreatLog_File_for_error (error_string)
    End If

    GetNotinGard
    Coloring
    Exit Sub

errHandler:
    error_string = error_string & vbCrLf & Err.Number & " - " & Err.Description
    Resume CleanExit
End Sub


'
'
Sub FillItemNew66()
On Error Resume Next
  error_string = ""
If txtFile.text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
    Dim currentvalue As String
    Dim Name As String
    Dim mStoreName As String
    Dim mPrice As Double
    Dim ExpireDate As String
    Dim itemcode As String
    Dim itemqty As Double
    Dim des As String
    Dim UnitName As String
    Dim DebitValue As String
    Dim CreditValue As String
    Dim s As String
    Dim rsDummyStore As ADODB.Recordset
  
  
      Dim sizename  As String
      Dim colorname  As String
      Dim mWidth  As String
      Dim mLength  As String
      Dim mHeight  As String
      Dim mArea  As String
      

  
    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")
    ExcelObj.Workbooks.Open txtFile.text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 
    With ExcelSheet
    i = 2
    
    Do Until .cells(i, 1) & "" = "" Or .cells(i, 2) & "" = ""
    
'    mStoreName = .cells(i, 1)
'    If i = 2 Then
'        s = "SELECT StoreID  FROM TblStore WHERE StoreName LIKE  '%" & Trim(mStoreName) & "%'"
'        Set rsDummyStore = New ADODB.Recordset
'        rsDummyStore.Open s, Cn, adOpenKeyset, adLockReadOnly
'        If Not rsDummyStore.EOF Then
'            DCboStoreName.BoundText = val(rsDummyStore!StoreId & "")
'        End If
'        rsDummyStore.Close
'    End If
    itemcode = .cells(i, 1)
    Name = .cells(i, 2)
    
     
    
    itemqty = .cells(i, 3)
    
   ' mPrice = .cells(i, 4)
    UnitName = .cells(i, 4)
        ExpireDate = .cells(i, 5)
        
        sizename = Trim(.cells(i, 6))
        colorname = Trim(.cells(i, 7))
        mWidth = Trim(.cells(i, 8))
        mLength = Trim(.cells(i, 9))
        mHeight = Trim(.cells(i, 10))
        mArea = Trim(.cells(i, 11))
        
            
            If itemcode = "D1010103001" Then
                itemcode = "D1010103001"
            End If
        
        addrow2 itemcode, itemqty, UnitName, mPrice, ExpireDate, sizename, colorname, mWidth, mLength, mHeight, mArea
       i = i + 1
       NewGrid.CountItems
    Loop
        End With
    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing

        If error_string <> "" Then
            CreatLog_File_for_error (error_string)
       End If
GetNotinGard
Coloring
End Sub


Private Sub Command2_Click()
  
    If chkNew.value = vbChecked Then
        FillItemNew
    Else
        If SystemOptions.WorkWithItemsDetails = True Then
            FillItemDete
        Else
            FillItem
        End If
    End If
End Sub
Public Sub CreatLog_File_for_error(str As String)
    Dim StrLogFileName As String
    Dim IntFreeFile As Integer
    Dim ss As String

    StrLogFileName = App.path & "\Gard.txt"

    If Dir(StrLogFileName) <> "" Then
        Kill StrLogFileName
    End If

    ss = "»Ì«‰ »«”„«¡  «·«’‰«› €Ì— «·„ÊÃÊœ… "
    ss = ss & vbCrLf & "Byte Informations Systems "
    ss = ss & vbCrLf & "BYTE "
    ss = ss & vbCrLf & "Create Date:- " & Now
    ss = ss & vbCrLf & str & vbCrLf
    IntFreeFile = FreeFile

    Open StrLogFileName For Output As #IntFreeFile
    Print #IntFreeFile, ss
    Close #IntFreeFile
End Sub

Private Sub Command3_Click()
Dim sql As String
 
sql = "delete dbo.Transaction_Details"
sql = sql & "  Where (AutoDetect = 1) and Transaction_ID=" & val(XPTxtBillID.text)

Cn.Execute sql
rs.Resync adAffectCurrent
Retrive
MsgBox " „ «·Õ–›", vbCritical
 End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 12
        FrmItemSearch.show vbModal
    End If

End Sub
 Function Retrive_Items_data1()
    Dim StrSQL  As String
    Dim row_count As Long
    Dim Num As Long
    Dim i As Long
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    StrSQL = "select * from TblItems where ItemID in(" & TxtItemsIDes.text & ")"
    rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If rs2.RecordCount > 0 Then
        
        If FG.TextMatrix(FG.rows - 1, FG.ColIndex("Code")) = "" Then
      FG.rows = FG.rows - 1
        End If
     With FG
     row_count = FG.rows
       rs2.MoveFirst
       .rows = rs2.RecordCount + .rows
        For Num = row_count To .rows - 1 'RsDetails.RecordCount
        .TextMatrix(Num, .ColIndex("Code")) = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
    
       'TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(Rs2("ItemName").value), "", Rs2("ItemName").value)
        rs2.MoveNext
        Next Num
        For i = row_count To .rows - 1 'RsDetails.RecordCount
          NewGrid.Grid_AfterEdit i, .ColIndex("Code")
        Next i
        NewGrid.Grid_AfterEdit row_count, .ColIndex("Code")
    End With
    End If
End Function
Private Sub DCboItemsName_KeyUp(KeyCode As Integer, Shift As Integer)
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

End Sub

Private Sub DCboStoreName_Click(Area As Integer)
    On Error Resume Next

    DCboStoreName_Change

End Sub

Private Sub DCboStoreName_Validate(Cancel As Boolean)
WriteDev
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
'    My_SQL = "  select branch_id,branch_name from TblBranchesData   "
'    fill_combo dcBranch, My_SQL
 
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
    Set NewGrid.VatGrid = Me.VatGrid
    Set NewGrid.TxtValueAdded = TxtValueAdded
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.Grid = FG
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.TxtItemCodeB = TxtItemCodeB
    
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

    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetStores Me.DCboStoreName
Dcombos.GetItemSGroups Me.XPCboGroup, False
        Dcombos.GetBranches Me.dcBranch
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboStoreName
    NewGrid.FillGrid
    StrSQL = "Select * From Transactions where Transaction_Type=30"
    StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    XPBtnMove_Click 2

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
    
   If DCboStoreName.text = "" Then Exit Sub
    Dim Account_Code_dynamic As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

        Account_Code_dynamic = get_store_Account(val(DCboStoreName.BoundText), "Account_Code")

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
                         Optional ToDate As Date, Optional myindex As Integer)
    Dim StrSQL As String
    Dim RsDetails As New ADODB.Recordset
    Dim RowNum As Integer
    Dim LngNoteID As Long

    On Error GoTo ErrTrap
 
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    If myindex = 1 Or myindex = 10 Then
    StrSQL = "SELECT    SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.Transaction_Details.Item_ID AS ItemID, "
    Else
    StrSQL = "SELECT  ExpiryDate  ,SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.Transaction_Details.Item_ID AS ItemID, "
    End If
    StrSQL = StrSQL & "  dbo.Transactions.StoreID,TblUnites.UnitName,TblUnites.UnitId, dbo.TblStore.StoreName, 1, dbo.TblItems.ItemCode,"
    'Transaction_Details.Height ,Transaction_Details.Length,Transaction_Details.Width,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemName, 1, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ClassId,"
    StrSQL = StrSQL & "  dbo.TblItemsclasses.SizeName AS ClassName, dbo.TblItemsSizes.SizeName AS SizeName, dbo.TblItemsColors.ColorName"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN "
    StrSQL = StrSQL & "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID INNER JOIN "
    StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"

    StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"
    StrSQL = StrSQL + " where   TblItems.ItemType=0 AND   dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & ""
    StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & ""
  '  If myindex <> 10 Then
        StrSQL = StrSQL & "   AND (dbo.Transactions.StoreID =" & StoreID & ")"
  '  End If
    If myindex = 10 Then
        'StrSQL = StrSQL & " and  Transaction_Details.Item_ID In (Select Item_ID from Transaction_Details TT  where TT.Price > 9000)       "
        StrSQL = StrSQL & " and  Transaction_Details.Item_ID In (Select Item_ID from Transaction_Details TT  where Item_Id  =498)       "
    End If
    StrSQL = StrSQL & "  GROUP BY dbo.Transaction_Details.Item_ID,TblUnites.UnitName,TblUnites.UnitId, dbo.Transactions.StoreID,dbo.TblStore.StoreName,"
    'Transaction_Details.Height ,Transaction_Details.Length,Transaction_Details.Width,
   
    StrSQL = StrSQL & "  dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize,"
    StrSQL = StrSQL & "   dbo.Transaction_Details.ClassId , dbo.TblItemsclasses.SizeName, dbo.TblItemsSizes.SizeName, dbo.TblItemsColors.ColorName"
    If myindex = 4 Then
    StrSQL = StrSQL & ",ExpiryDate "
    End If
    
  StrSQL = ""
StrSQL = StrSQL & "WITH LastStockTaking AS ("
StrSQL = StrSQL & " SELECT TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId, MAX(T.Transaction_Date) AS LastStockDate"
StrSQL = StrSQL & " FROM dbo.Transactions T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " WHERE T.Transaction_Type = 30"
StrSQL = StrSQL & " AND T.Transaction_Date >= " & SQLDate(FromDate, True)
StrSQL = StrSQL & " AND T.Transaction_Date <= " & SQLDate(ToDate, True)
StrSQL = StrSQL & " AND T.StoreID = " & StoreID
StrSQL = StrSQL & " GROUP BY TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId"
StrSQL = StrSQL & "),"
StrSQL = StrSQL & " InitialStock AS ("
StrSQL = StrSQL & " SELECT TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId, SUM(TD.Quantity) AS QtyStockTake"
StrSQL = StrSQL & " FROM dbo.Transactions T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " WHERE T.Transaction_Type = 30"
StrSQL = StrSQL & " AND T.Transaction_Date >= " & SQLDate(FromDate, True)
StrSQL = StrSQL & " AND T.Transaction_Date <= " & SQLDate(ToDate, True)
StrSQL = StrSQL & " AND T.StoreID = " & StoreID
StrSQL = StrSQL & " AND T.Transaction_Date IN ("
StrSQL = StrSQL & "     SELECT MAX(T2.Transaction_Date)"
StrSQL = StrSQL & "     FROM dbo.Transactions T2"
StrSQL = StrSQL & "     INNER JOIN dbo.Transaction_Details TD2 ON T2.Transaction_ID = TD2.Transaction_ID"
StrSQL = StrSQL & "     WHERE T2.Transaction_Type = 30"
StrSQL = StrSQL & "     AND T2.Transaction_Date >= " & SQLDate(FromDate, True)
StrSQL = StrSQL & "     AND T2.Transaction_Date <= " & SQLDate(ToDate, True)
StrSQL = StrSQL & "     AND T2.StoreID = T.StoreID"
StrSQL = StrSQL & "     AND TD2.Item_ID = TD.Item_ID"
StrSQL = StrSQL & "     AND TD2.UnitId = TD.UnitId"
StrSQL = StrSQL & "     AND TD2.ColorID = TD.ColorID"
StrSQL = StrSQL & "     AND TD2.ItemSize = TD.ItemSize"
StrSQL = StrSQL & "     AND TD2.ClassId = TD.ClassId"
StrSQL = StrSQL & " )"
StrSQL = StrSQL & " GROUP BY TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId"
StrSQL = StrSQL & "),"
StrSQL = StrSQL & " StockMovements AS ("
StrSQL = StrSQL & " SELECT TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId,"
StrSQL = StrSQL & "        SUM(TD.Quantity * TT.StockEffect) AS MovementsAfterStockTake"
StrSQL = StrSQL & " FROM dbo.Transactions T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " INNER JOIN dbo.TransactionTypes TT ON T.Transaction_Type = TT.Transaction_Type"
StrSQL = StrSQL & " LEFT JOIN LastStockTaking LST ON LST.Item_ID = TD.Item_ID AND LST.StoreID = T.StoreID"
StrSQL = StrSQL & "   AND LST.UnitId = TD.UnitId"
StrSQL = StrSQL & "   AND ISNULL(LST.ColorID, 1) = ISNULL(TD.ColorID, 1)"
StrSQL = StrSQL & "   AND ISNULL(LST.ItemSize, 1) = ISNULL(TD.ItemSize, 1)"
StrSQL = StrSQL & "   AND ISNULL(LST.ClassId, 1) = ISNULL(TD.ClassId, 1)"
StrSQL = StrSQL & " WHERE T.Transaction_Date > ISNULL(LST.LastStockDate, " & SQLDate(FromDate, True) & ")"
StrSQL = StrSQL & " AND T.Transaction_Date <= " & SQLDate(ToDate, True)
StrSQL = StrSQL & " AND T.StoreID = " & StoreID
StrSQL = StrSQL & " AND T.Transaction_Type <> 30"
StrSQL = StrSQL & " GROUP BY TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId"
StrSQL = StrSQL & "),"
StrSQL = StrSQL & " ItemsAll AS ("
StrSQL = StrSQL & " SELECT DISTINCT TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId"
StrSQL = StrSQL & " FROM dbo.Transactions T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " WHERE T.Transaction_Date >= " & SQLDate(FromDate, True)
StrSQL = StrSQL & " AND T.Transaction_Date <= " & SQLDate(ToDate, True)
StrSQL = StrSQL & " AND T.StoreID = " & StoreID
StrSQL = StrSQL & ")"
StrSQL = StrSQL & " SELECT"
StrSQL = StrSQL & "   ia.Item_ID AS ItemID,"
StrSQL = StrSQL & "   ia.StoreID,"
StrSQL = StrSQL & "   ia.UnitId,"
StrSQL = StrSQL & "   ia.ColorID,"
StrSQL = StrSQL & "   ia.ItemSize,"
StrSQL = StrSQL & "   ia.ClassId,"
StrSQL = StrSQL & "   TblUnites.UnitName,"
StrSQL = StrSQL & "   TblUnites.UnitID,"
StrSQL = StrSQL & "   TblStore.StoreName,"
StrSQL = StrSQL & "   TblItems.ItemCode,"
StrSQL = StrSQL & "   TblItems.ItemName,"
StrSQL = StrSQL & "   TblItemsclasses.SizeName AS ClassName,"
StrSQL = StrSQL & "   TblItemsSizes.SizeName AS SizeName,"
StrSQL = StrSQL & "   TblItemsColors.ColorName,"
StrSQL = StrSQL & "   ISNULL(i.QtyStockTake, 0) AS InitialStock,"
StrSQL = StrSQL & "   ISNULL(sm.MovementsAfterStockTake, 0) AS MovementsAfterStockTake,"
StrSQL = StrSQL & "   (ISNULL(i.QtyStockTake, 0) + ISNULL(sm.MovementsAfterStockTake, 0)) AS NetStock"
StrSQL = StrSQL & " FROM ItemsAll ia"
StrSQL = StrSQL & " LEFT JOIN InitialStock i ON ia.Item_ID = i.Item_ID AND ia.StoreID = i.StoreID"
StrSQL = StrSQL & "   AND ia.UnitId = i.UnitId"
StrSQL = StrSQL & "   AND ISNULL(ia.ColorID, 1) = ISNULL(i.ColorID, 1)"
StrSQL = StrSQL & "   AND ISNULL(ia.ItemSize, 1) = ISNULL(i.ItemSize, 1)"
StrSQL = StrSQL & "   AND ISNULL(ia.ClassId, 1) = ISNULL(i.ClassId, 1)"
StrSQL = StrSQL & " LEFT JOIN StockMovements sm ON ia.Item_ID = sm.Item_ID AND ia.StoreID = sm.StoreID"
StrSQL = StrSQL & "   AND ia.UnitId = sm.UnitId"
StrSQL = StrSQL & "   AND ISNULL(ia.ColorID, 1) = ISNULL(sm.ColorID, 1)"
StrSQL = StrSQL & "   AND ISNULL(ia.ItemSize, 1) = ISNULL(sm.ItemSize, 1)"
StrSQL = StrSQL & "   AND ISNULL(ia.ClassId, 1) = ISNULL(sm.ClassId, 1)"
StrSQL = StrSQL & " LEFT JOIN dbo.TblItems ON ia.Item_ID = TblItems.ItemID"
StrSQL = StrSQL & " LEFT JOIN dbo.TblUnites ON ia.UnitId = TblUnites.UnitID"
StrSQL = StrSQL & " LEFT JOIN dbo.TblStore ON ia.StoreID = TblStore.StoreID"
StrSQL = StrSQL & " LEFT JOIN dbo.TblItemsSizes ON ia.ItemSize = TblItemsSizes.SizeId"
StrSQL = StrSQL & " LEFT JOIN dbo.TblItemsclasses ON ia.ClassId = TblItemsclasses.SizeId"
StrSQL = StrSQL & " LEFT JOIN dbo.TblItemsColors ON ia.ColorID = TblItemsColors.ColorID"
StrSQL = StrSQL & " WHERE TblItems.ItemType = 0"
StrSQL = StrSQL & " AND (ISNULL(i.QtyStockTake, 0) + ISNULL(sm.MovementsAfterStockTake, 0)) <> 0"
StrSQL = StrSQL & " ORDER BY TblItems.ItemCode, TblUnites.UnitID, ia.ColorID, ia.ItemSize, ia.ClassId"

    StrSQL = ""
StrSQL = StrSQL & "WITH LastStockTaking AS ("
StrSQL = StrSQL & " SELECT"
StrSQL = StrSQL & "   TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId,"
StrSQL = StrSQL & "   T.Transaction_Date, T.Transaction_ID,"
StrSQL = StrSQL & "   ROW_NUMBER() OVER (PARTITION BY "
StrSQL = StrSQL & "     TD.Item_ID, T.StoreID, TD.UnitId,"
StrSQL = StrSQL & "     ISNULL(TD.ColorID,-1), ISNULL(TD.ItemSize,-1), ISNULL(TD.ClassId,-1)"
StrSQL = StrSQL & "     ORDER BY T.Transaction_Date DESC, T.Transaction_ID DESC"
StrSQL = StrSQL & "   ) AS RN"
StrSQL = StrSQL & " FROM dbo.Transactions T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " WHERE T.Transaction_Type = 30"
StrSQL = StrSQL & " AND T.Transaction_Date >= " & SQLDate(FromDate, True)
StrSQL = StrSQL & " AND T.Transaction_Date <= " & SQLDate(ToDate, True)
StrSQL = StrSQL & " AND T.StoreID = " & StoreID
StrSQL = StrSQL & "),"

StrSQL = StrSQL & " LastStock AS ("
StrSQL = StrSQL & " SELECT"
StrSQL = StrSQL & "   Item_ID, StoreID, UnitId, ColorID, ItemSize, ClassId,"
StrSQL = StrSQL & "   Transaction_Date AS LastStockDate,"
StrSQL = StrSQL & "   Transaction_ID AS LastStockTransId"
StrSQL = StrSQL & " FROM LastStockTaking"
StrSQL = StrSQL & " WHERE RN = 1"
StrSQL = StrSQL & "),"

StrSQL = StrSQL & " InitialStock AS ("
StrSQL = StrSQL & " SELECT"
StrSQL = StrSQL & "   TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId,"
StrSQL = StrSQL & "   SUM(TD.Quantity) AS QtyStockTake"
StrSQL = StrSQL & " FROM dbo.Transactions T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " INNER JOIN LastStock LS ON"
StrSQL = StrSQL & "   LS.Item_ID = TD.Item_ID"
StrSQL = StrSQL & "   AND LS.StoreID = T.StoreID"
StrSQL = StrSQL & "   AND LS.UnitId = TD.UnitId"
StrSQL = StrSQL & "   AND ISNULL(LS.ColorID,-1) = ISNULL(TD.ColorID,-1)"
StrSQL = StrSQL & "   AND ISNULL(LS.ItemSize,-1) = ISNULL(TD.ItemSize,-1)"
StrSQL = StrSQL & "   AND ISNULL(LS.ClassId,-1) = ISNULL(TD.ClassId,-1)"
StrSQL = StrSQL & "   AND LS.LastStockTransId = T.Transaction_ID"
StrSQL = StrSQL & " WHERE T.Transaction_Type = 30"
StrSQL = StrSQL & " GROUP BY TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId"
StrSQL = StrSQL & "),"

StrSQL = StrSQL & " StockMovements AS ("
StrSQL = StrSQL & " SELECT"
StrSQL = StrSQL & "   TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId,"
StrSQL = StrSQL & "   SUM(TD.Quantity * TT.StockEffect) AS MovementsAfterStockTake"
StrSQL = StrSQL & " FROM dbo.Transactions T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " INNER JOIN dbo.TransactionTypes TT ON T.Transaction_Type = TT.Transaction_Type"
StrSQL = StrSQL & " LEFT JOIN LastStock LS ON"
StrSQL = StrSQL & "   LS.Item_ID = TD.Item_ID"
StrSQL = StrSQL & "   AND LS.StoreID = T.StoreID"
StrSQL = StrSQL & "   AND LS.UnitId = TD.UnitId"
StrSQL = StrSQL & "   AND ISNULL(LS.ColorID,-1) = ISNULL(TD.ColorID,-1)"
StrSQL = StrSQL & "   AND ISNULL(LS.ItemSize,-1) = ISNULL(TD.ItemSize,-1)"
StrSQL = StrSQL & "   AND ISNULL(LS.ClassId,-1) = ISNULL(TD.ClassId,-1)"
StrSQL = StrSQL & " WHERE T.StoreID = " & StoreID
StrSQL = StrSQL & " AND T.Transaction_Type <> 30"
StrSQL = StrSQL & " AND ("
StrSQL = StrSQL & "   T.Transaction_Date > ISNULL(LS.LastStockDate, " & SQLDate(FromDate, True) & ")"
StrSQL = StrSQL & "   OR (T.Transaction_Date = ISNULL(LS.LastStockDate, " & SQLDate(FromDate, True) & ")"
StrSQL = StrSQL & "       AND T.Transaction_ID > ISNULL(LS.LastStockTransId,0))"
StrSQL = StrSQL & " )"
StrSQL = StrSQL & " AND T.Transaction_Date <= " & SQLDate(ToDate, True)
StrSQL = StrSQL & " GROUP BY TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId"
StrSQL = StrSQL & "),"

StrSQL = StrSQL & " ItemsAll AS ("
StrSQL = StrSQL & " SELECT DISTINCT"
StrSQL = StrSQL & "   TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId"
StrSQL = StrSQL & " FROM dbo.Transactions T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " WHERE T.Transaction_Date >= " & SQLDate(FromDate, True)
StrSQL = StrSQL & " AND T.Transaction_Date <= " & SQLDate(ToDate, True)
StrSQL = StrSQL & " AND T.StoreID = " & StoreID
StrSQL = StrSQL & ")"

StrSQL = StrSQL & " SELECT"
StrSQL = StrSQL & "   ia.Item_ID AS ItemID, ia.StoreID, ia.UnitId, ia.ColorID, ia.ItemSize, ia.ClassId,"
StrSQL = StrSQL & "   TblUnites.UnitName, TblUnites.UnitID, TblStore.StoreName,"
StrSQL = StrSQL & "   TblItems.ItemCode, TblItems.ItemName,"
StrSQL = StrSQL & "   TblItemsclasses.SizeName AS ClassName,"
StrSQL = StrSQL & "   TblItemsSizes.SizeName AS SizeName,"
StrSQL = StrSQL & "   TblItemsColors.ColorName,"
StrSQL = StrSQL & "   ISNULL(i.QtyStockTake,0) AS InitialStock,"
StrSQL = StrSQL & "   ISNULL(sm.MovementsAfterStockTake,0) AS MovementsAfterStockTake,"
StrSQL = StrSQL & "   ISNULL(i.QtyStockTake,0) + ISNULL(sm.MovementsAfterStockTake,0) AS NetStock"
StrSQL = StrSQL & " FROM ItemsAll ia"
StrSQL = StrSQL & " LEFT JOIN InitialStock i ON"
StrSQL = StrSQL & "   ia.Item_ID = i.Item_ID AND ia.StoreID = i.StoreID AND ia.UnitId = i.UnitId"
StrSQL = StrSQL & "   AND ISNULL(ia.ColorID,-1) = ISNULL(i.ColorID,-1)"
StrSQL = StrSQL & "   AND ISNULL(ia.ItemSize,-1) = ISNULL(i.ItemSize,-1)"
StrSQL = StrSQL & "   AND ISNULL(ia.ClassId,-1) = ISNULL(i.ClassId,-1)"
StrSQL = StrSQL & " LEFT JOIN StockMovements sm ON"
StrSQL = StrSQL & "   ia.Item_ID = sm.Item_ID AND ia.StoreID = sm.StoreID AND ia.UnitId = sm.UnitId"
StrSQL = StrSQL & "   AND ISNULL(ia.ColorID,-1) = ISNULL(sm.ColorID,-1)"
StrSQL = StrSQL & "   AND ISNULL(ia.ItemSize,-1) = ISNULL(sm.ItemSize,-1)"
StrSQL = StrSQL & "   AND ISNULL(ia.ClassId,-1) = ISNULL(sm.ClassId,-1)"
StrSQL = StrSQL & " LEFT JOIN TblItems ON ia.Item_ID = TblItems.ItemID"
StrSQL = StrSQL & " LEFT JOIN TblUnites ON ia.UnitId = TblUnites.UnitID"
StrSQL = StrSQL & " LEFT JOIN TblStore ON ia.StoreID = TblStore.StoreID"
StrSQL = StrSQL & " LEFT JOIN TblItemsSizes ON ia.ItemSize = TblItemsSizes.SizeId"
StrSQL = StrSQL & " LEFT JOIN TblItemsclasses ON ia.ClassId = TblItemsclasses.SizeId"
StrSQL = StrSQL & " LEFT JOIN TblItemsColors ON ia.ColorID = TblItemsColors.ColorID"
StrSQL = StrSQL & " WHERE TblItems.ItemType = 0"
StrSQL = StrSQL & " AND (ISNULL(i.QtyStockTake,0) + ISNULL(sm.MovementsAfterStockTake,0)) <> 0"
StrSQL = StrSQL & " ORDER BY TblItems.ItemCode, TblUnites.UnitID, ia.ColorID, ia.ItemSize, ia.ClassId"

    '===========================================
' Stock Balance Query (Matches Item Card)
' Same original CTE names:
'   LastStockTaking / InitialStock / StockMovements / ItemsAll
' InitialStock  = Opening balance before FromDate (SUM StockEffect*Qty)
' Movements     = SUM StockEffect*Qty within period
' NetStock      = InitialStock + Movements
'===========================================

StrSQL = ""

StrSQL = StrSQL & "WITH LastStockTaking AS ("
StrSQL = StrSQL & " SELECT TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId, MAX(T.Transaction_Date) AS LastStockDate"
StrSQL = StrSQL & " FROM dbo.Transactions T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " WHERE T.Transaction_Type = 30"
StrSQL = StrSQL & "   AND T.StoreID = " & StoreID
StrSQL = StrSQL & "   AND T.Transaction_Date <= " & SQLDate(ToDate, True)
StrSQL = StrSQL & " GROUP BY TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId"
StrSQL = StrSQL & "),"

StrSQL = StrSQL & " InitialStock AS ("
StrSQL = StrSQL & " SELECT TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId,"
StrSQL = StrSQL & "        SUM(ISNULL(TT.StockEffect * TD.Quantity, 0)) AS QtyStockTake"
StrSQL = StrSQL & " FROM dbo.Transactions T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " INNER JOIN dbo.TransactionTypes TT ON T.Transaction_Type = TT.Transaction_Type"
StrSQL = StrSQL & " WHERE TT.StockEffect <> 0"
StrSQL = StrSQL & "   AND T.StoreID = " & StoreID
StrSQL = StrSQL & "   AND T.Transaction_Date < " & SQLDate(FromDate, True)
StrSQL = StrSQL & "   AND T.Transaction_Date <= " & SQLDate(ToDate, True)
StrSQL = StrSQL & " GROUP BY TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId"
StrSQL = StrSQL & "),"

StrSQL = StrSQL & " StockMovements AS ("
StrSQL = StrSQL & " SELECT TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId,"
StrSQL = StrSQL & "        SUM(ISNULL(TT.StockEffect * TD.Quantity, 0)) AS MovementsAfterStockTake"
StrSQL = StrSQL & " FROM dbo.Transactions T"
StrSQL = StrSQL & " INNER JOIN dbo.Transaction_Details TD ON T.Transaction_ID = TD.Transaction_ID"
StrSQL = StrSQL & " INNER JOIN dbo.TransactionTypes TT ON T.Transaction_Type = TT.Transaction_Type"
StrSQL = StrSQL & " WHERE TT.StockEffect <> 0"
StrSQL = StrSQL & "   AND T.StoreID = " & StoreID
StrSQL = StrSQL & "   AND T.Transaction_Date >= " & SQLDate(FromDate, True)
StrSQL = StrSQL & "   AND T.Transaction_Date <= " & SQLDate(ToDate, True)
StrSQL = StrSQL & " GROUP BY TD.Item_ID, T.StoreID, TD.UnitId, TD.ColorID, TD.ItemSize, TD.ClassId"
StrSQL = StrSQL & "),"

StrSQL = StrSQL & " ItemsAll AS ("
StrSQL = StrSQL & " SELECT DISTINCT x.Item_ID, x.StoreID, x.UnitId, x.ColorID, x.ItemSize, x.ClassId"
StrSQL = StrSQL & " FROM ("
StrSQL = StrSQL & "   SELECT Item_ID, StoreID, UnitId, ColorID, ItemSize, ClassId FROM InitialStock"
StrSQL = StrSQL & "   UNION ALL"
StrSQL = StrSQL & "   SELECT Item_ID, StoreID, UnitId, ColorID, ItemSize, ClassId FROM StockMovements"
StrSQL = StrSQL & " ) x"
StrSQL = StrSQL & ")"

StrSQL = StrSQL & " SELECT"
StrSQL = StrSQL & "   ia.Item_ID AS ItemID,"
StrSQL = StrSQL & "   ia.StoreID,"
StrSQL = StrSQL & "   ia.UnitId,"
StrSQL = StrSQL & "   ia.ColorID,"
StrSQL = StrSQL & "   ia.ItemSize,"
StrSQL = StrSQL & "   ia.ClassId,"
StrSQL = StrSQL & "   u.UnitName,"
StrSQL = StrSQL & "   u.UnitID,"
StrSQL = StrSQL & "   s.StoreName,"
StrSQL = StrSQL & "   it.ItemCode,"
StrSQL = StrSQL & "   it.ItemName,"
StrSQL = StrSQL & "   cls.SizeName AS ClassName,"
StrSQL = StrSQL & "   sz.SizeName AS SizeName,"
StrSQL = StrSQL & "   c.ColorName,"
StrSQL = StrSQL & "   ISNULL(i.QtyStockTake, 0) AS InitialStock,"
StrSQL = StrSQL & "   ISNULL(sm.MovementsAfterStockTake, 0) AS MovementsAfterStockTake,"
StrSQL = StrSQL & "   (ISNULL(i.QtyStockTake, 0) + ISNULL(sm.MovementsAfterStockTake, 0)) AS NetStock"
StrSQL = StrSQL & " FROM ItemsAll ia"
StrSQL = StrSQL & " LEFT JOIN InitialStock i"
StrSQL = StrSQL & "   ON ia.Item_ID = i.Item_ID AND ia.StoreID = i.StoreID AND ia.UnitId = i.UnitId"
StrSQL = StrSQL & "  AND ISNULL(ia.ColorID, -1) = ISNULL(i.ColorID, -1)"
StrSQL = StrSQL & "  AND ISNULL(ia.ItemSize, -1) = ISNULL(i.ItemSize, -1)"
StrSQL = StrSQL & "  AND ISNULL(ia.ClassId, -1) = ISNULL(i.ClassId, -1)"
StrSQL = StrSQL & " LEFT JOIN StockMovements sm"
StrSQL = StrSQL & "   ON ia.Item_ID = sm.Item_ID AND ia.StoreID = sm.StoreID AND ia.UnitId = sm.UnitId"
StrSQL = StrSQL & "  AND ISNULL(ia.ColorID, -1) = ISNULL(sm.ColorID, -1)"
StrSQL = StrSQL & "  AND ISNULL(ia.ItemSize, -1) = ISNULL(sm.ItemSize, -1)"
StrSQL = StrSQL & "  AND ISNULL(ia.ClassId, -1) = ISNULL(sm.ClassId, -1)"
StrSQL = StrSQL & " LEFT JOIN dbo.TblItems it ON ia.Item_ID = it.ItemID"
StrSQL = StrSQL & " LEFT JOIN dbo.TblUnites u ON ia.UnitId = u.UnitID"
StrSQL = StrSQL & " LEFT JOIN dbo.TblStore s ON ia.StoreID = s.StoreID"
StrSQL = StrSQL & " LEFT JOIN dbo.TblItemsSizes sz ON ia.ItemSize = sz.SizeId"
StrSQL = StrSQL & " LEFT JOIN dbo.TblItemsclasses cls ON ia.ClassId = cls.SizeId"
StrSQL = StrSQL & " LEFT JOIN dbo.TblItemsColors c ON ia.ColorID = c.ColorID"
StrSQL = StrSQL & " WHERE it.ItemType = 0"
StrSQL = StrSQL & "   AND (ISNULL(i.QtyStockTake, 0) + ISNULL(sm.MovementsAfterStockTake, 0)) <> 0"
StrSQL = StrSQL & " ORDER BY it.ItemCode, u.UnitID, ia.ColorID, ia.ItemSize, ia.ClassId"

    Dim LngItemID As Long
    Dim LngUnitID As Long
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For RowNum = 1 To RsDetails.RecordCount

            With FG
                
                
               
              If myindex = 4 Then
             .TextMatrix(RowNum, .ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
             End If
                
                .TextMatrix(RowNum, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID").value), "", RsDetails("ItemID").value)
                .TextMatrix(RowNum, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID").value), "", RsDetails("ItemID").value)
                
               ' .TextMatrix(RowNum, FG.ColIndex("Length")) = IIf(IsNull(RsDetails("Length").value), "", RsDetails("Length").value)
               ' .TextMatrix(RowNum, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height").value), "", RsDetails("Height").value)
               ' .TextMatrix(RowNum, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width").value), "", RsDetails("Width").value)
                
                '          .TextMatrix(RowNum, FG.ColIndex("Count1")) = IIf(IsNull(RsDetails("SUMQTY").value), "", RsDetails("SUMQTY").value)
             
                 Dim UnitName As String
        Dim unotfactor As Double
        
   'GetDefaultItemUnit val(.TextMatrix(RowNum, FG.ColIndex("Code"))), LngUnitID, UnitName, unotfactor, val(.TextMatrix(RowNum, FG.ColIndex("Code")))
       
       
'             FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = LngUnitID ' IIf(IsNull(RsDetails("UnitID")), 1, (RsDetails("UnitID").value))
'              FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = UnitName ' IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
 
             FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), 1, (RsDetails("UnitID").value))
              FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
 
 
                LngItemID = val(.TextMatrix(RowNum, .ColIndex("Code")))
                LngUnitID = FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))
                  If LngItemID = 6396 Then
                       LngItemID = 6396
                  End If
                  
        If SystemOptions.CostStarting = True Then
                    Dim FirstPeriodDateInthisYear  As Date
                    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

                fromcostdate = DateAdd("d", -1, FirstPeriodDateInthisYear)
                fromcostdate = Replace(Format$(fromcostdate, "MM/DD/yyyy"), "-", "/")
                .TextMatrix(RowNum, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , fromcostdate, DTPickerAccTo, , LngUnitID)
          Else
                .TextMatrix(RowNum, .ColIndex("Price")) = 0 ' ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , DTPickerAccTo, , LngUnitID)
                .TextMatrix(RowNum, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , DTPickerAccTo, , LngUnitID)
            End If


                
                    
                '.TextMatrix(RowNum, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price").value), "", RsDetails("Price").value)
                '        .TextMatrix(RowNum, FG.ColIndex("Valu")) = Val(.TextMatrix(RowNum, .ColIndex("Price"))) * Val(.TextMatrix(RowNum, .ColIndex("Count1")))
            End With

            FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            FG.TextMatrix(RowNum, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("NetStock")), 0, (RsDetails("NetStock").value))
          NewGrid.CountItems
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

Public Function Retrive2(Optional GroupID As Integer)
    Dim StrSQL As String
    Dim RsDetails As New ADODB.Recordset
    Dim RowNum As Integer
    Dim LngNoteID As Long

    On Error GoTo ErrTrap
 
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
     
StrSQL = "SELECT     dbo.TblItems.ItemID, dbo.TblItems.GroupID, dbo.TblItemsUnits.UnitID, dbo.TblItemsUnits.DefaultUnit, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
StrSQL = StrSQL & "   FROM         dbo.TblItems INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"

StrSQL = StrSQL & "  Where (dbo.TblItemsUnits.DefaultUnit = 1) And (dbo.TblItems.GroupID =" & GroupID & ")"
StrSQL = StrSQL & "  ORDER BY dbo.TblItems.Itemcode"
    
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
                FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), 1, (RsDetails("UnitID").value))
                FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
   
                LngItemID = val(.TextMatrix(RowNum, .ColIndex("Code")))
                LngUnitID = FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))
                  
                  
                   If SystemOptions.CostStarting = True Then
    Dim FirstPeriodDateInthisYear  As Date
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

     fromcostdate = DateAdd("d", -1, FirstPeriodDateInthisYear)
fromcostdate = Replace(Format$(fromcostdate, "MM/DD/yyyy"), "-", "/")
                .TextMatrix(RowNum, .ColIndex("Price")) = 0 'ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , fromcostdate, DTPickerAccTo, , LngUnitID)
          Else
          .TextMatrix(RowNum, .ColIndex("Price")) = 0 ' ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , DTPickerAccTo, , LngUnitID)
End If


                
                    
                '.TextMatrix(RowNum, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price").value), "", RsDetails("Price").value)
                '        .TextMatrix(RowNum, FG.ColIndex("Valu")) = Val(.TextMatrix(RowNum, .ColIndex("Price"))) * Val(.TextMatrix(RowNum, .ColIndex("Count1")))
            End With

            FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = 1 ' IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = 1 ' IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ClassID")) = 1 ' IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            FG.TextMatrix(RowNum, FG.ColIndex("Count")) = 0
          NewGrid.CountItems
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

Private Sub Opt_Click(Index As Integer)
If Index = 1 Or Index = 4 Then
retrive1 val(DCboStoreName.BoundText), DTPickerAccFrom.value, DTPickerAccTo.value, Index

End If
End Sub

Private Sub SearchCashCustomer_Click(Index As Integer)
 '   If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch2
        FrmItemSearch2.RetrunType = 4
         FrmItemSearch2.show vbModal
    'End If

End Sub

Private Sub txtStoreSearch_Validate(Cancel As Boolean)
Dim s As String
If Trim(txtStoreSearch) <> "" Then
    s = "Select * from TblStore Where StoreName Like '%" & Trim(txtStoreSearch) & "%'"
    Dim rsDummy As New ADODB.Recordset
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If Not rsDummy.EOF Then
        DCboStoreName.BoundText = val(rsDummy!StoreID & "")
        
    End If
End If
txtStoreSearch = ""
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
'   On Error GoTo ErrTrap
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
            
            Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
 
                DCboStoreName.Enabled = True
              '  TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore
            Else
                dcBranch.Enabled = True
 
                DCboStoreName.Enabled = True
 
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
'                TxtStoreID.Enabled = True
            End If
                    
                    
        

      If SystemOptions.usertype <> UserAdminAll Then
                            If checkmanyBranches = False Then
                                   Me.dcBranch.Enabled = True
                                   Else
                                    Me.dcBranch.Enabled = True
                             End If
                    
                      If checkmanyStores = False Then
                                   Me.DCboStoreName.Enabled = True
                                    
                                   Else
                                   Me.DCboStoreName.Enabled = True
 
                             End If
                                  
           End If



            Me.dcBranch.BoundText = Current_branch
            Dim FirstPeriodDateInthisYear  As Date
'getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
            getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            Me.DTPickerAccFrom = FirstPeriodDateInthisYear
            DTPickerAccTo.value = Date
            opt(2).value = True

        Case 1
                                     If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                  

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            'If AvailableDeal = True Then
            TxtModFlg.text = "E"
            DCboStoreName_Change
            Me.DCboUserName.BoundText = user_id

            'End If
        Case 2
                                     If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                  
            SaveData

        Case 3
            Call Undo

        Case 4
                             If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                  
                  
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_TransAction

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            printing

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
            FrmBalanceSearch.mIndex = 0
            FrmBalanceSearch.mTransaction_Type = 30
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

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBillID.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If True = True Then
                If Not rs.RecordCount < 1 Then
                    DeleteTransactiomsVoucher val(Text1.text)
                    DeleteTransactiomsVoucher val(Text2.text)
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

  '  On Error GoTo ErrTrap
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
        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If
GetNotinGard

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
 
        rs("Transaction_Type").value = 30
        rs("UserID").value = user_id
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, DCboStoreName.BoundText)
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
                RSTransDetails("Transaction_ID").value = XPTxtBillID.text
                
                RSTransDetails("AutoDetect").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("AutoDetect")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("AutoDetect"))))
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
                   RSTransDetails("ParrtNoCode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))))
         RSTransDetails("ItemDetailedCode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))))
'
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("Price").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
                RSTransDetails("Height").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Height")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Height"))))
                
                RSTransDetails("Width").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Width")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Width"))))
                RSTransDetails("length").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("length")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("length"))))
                RSTransDetails("Height").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Height")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Height"))))
                RSTransDetails("Area").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Area")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Area"))))
            
                RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
                ' IIf((FG.TextMatrix(RowNum, FG.ColIndex("BranchId")) = ""), 1, Val(FG.TextMatrix(RowNum, FG.ColIndex("BranchId"))))
               
                ' RSTransDetails("ItemSize").value = _
                  IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), "", Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))
                RSTransDetails("ProductionDate").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")))
                RSTransDetails("ExpiryDate").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")))
                
                    
                RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", 1, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
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
 Coloring
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
        rs.Find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
'Me.TxtModFlg.text = "R"
    Screen.MousePointer = vbArrowHourglass
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

    DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    StrSQL = "SELECT AutoDetect,    dbo.TblItems.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.GroupID, dbo.TblItems.HaveSerial, dbo.TblItems.LastUpdate, "
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
 StrSQL = StrSQL + "                      dbo.Transaction_Details.ItemDetailedCode ,ParrtNoCode,dbo.Transaction_Details.LotNO,Transaction_Details.ExpiryDate,"
 StrSQL = StrSQL + "                       Transaction_Details.Height,Transaction_Details.Width,Transaction_Details.length"
 StrSQL = StrSQL + " FROM         dbo.TblItems INNER JOIN"
 StrSQL = StrSQL + "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID LEFT OUTER JOIN"
 StrSQL = StrSQL + "                      dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"

'    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
        
    StrSQL = StrSQL + " order by Transaction_Details.id "

 
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For RowNum = 1 To RsDetails.RecordCount

            With FG
            
            .TextMatrix(RowNum, FG.ColIndex("AutoDetect")) = IIf(IsNull(RsDetails("AutoDetect").value), 0, RsDetails("AutoDetect").value)
            
                .TextMatrix(RowNum, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID").value), "", RsDetails("Item_ID").value)
                .TextMatrix(RowNum, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID").value), "", RsDetails("Item_ID").value)
                .TextMatrix(RowNum, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Showqty").value), "", RsDetails("Showqty").value)
                .TextMatrix(RowNum, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial").value), "", RsDetails("ItemSerial").value)
                .TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = IIf(IsNull(RsDetails("HaveSerial").value), "", RsDetails("HaveSerial").value)
                FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
                   FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            
                If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                    FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
                Else
                    FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
                End If

                .TextMatrix(RowNum, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice").value), "", RsDetails("ShowPrice").value)
                .TextMatrix(RowNum, FG.ColIndex("Valu")) = val(.TextMatrix(RowNum, .ColIndex("Price"))) * val(.TextMatrix(RowNum, .ColIndex("Count")))
            End With

            FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(RowNum, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            
            FG.TextMatrix(RowNum, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))

            FG.TextMatrix(RowNum, FG.ColIndex("length")) = IIf(IsNull(RsDetails("length")), "", (RsDetails("length").value))
            FG.TextMatrix(RowNum, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))
            
        
            FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
      FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) = IIf(IsNull(RsDetails("ParrtNoCode")), "", (RsDetails("ParrtNoCode").value))
FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode")) = IIf(IsNull(RsDetails("ItemDetailedCode")), "", (RsDetails("ItemDetailedCode").value))

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
    Coloring
    NewGrid.CountItems
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub printing()
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Set BalanceReport = New ClsOpeningBalanceReport
        BalanceReport.ShowOpeningBalanceData XPTxtBillID.text, , 1
    End If

    Exit Sub
ErrTrap:
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
Private Function AvailableDeal() As Boolean
    Dim RowNum As Integer
    Dim Msg As String
    Dim StrSQL As String
    
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
 Label1.Caption = "Actual Page No."
' Ele(2).Caption = "P"
lbl(12).Caption = "BarCode"
    Ele(2).Caption = "Period"
    lbl(10).Caption = "From"
    lbl(11).Caption = "To"
    Frame1.Caption = "Input Methos"
    opt(0).Caption = "File"
    opt(1).Caption = "All Items"
    opt(2).Caption = "Manual"
    Command1.Caption = "Browse...."
    Me.Caption = "Actual Inventory"
    C1Elastic6.Caption = Me.Caption

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
   
    With Me.FG
        '.TextMatrix(0, .ColIndex("NewItem")) = "NewItem"
    End With
   
    'NewItem

End Sub

Function GetNotinGard()
Dim RowNum As Double
Dim ItemString As String
Dim StrSQL As String
    Dim RsDetails As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngUnitID As Long
  Dim ColorID As Integer
   Dim sizeid As Integer
    Dim ClassId As Integer
    Dim ParrtNoCode As String
    Dim ItemDetailedCode As String
'Exit Function
Dim Price As Double

ItemString = ""
With Me.FG
        For RowNum = 1 To FG.rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            
             If SystemOptions.WorkWithItemsDetails = True Then
             
                         If FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) <> "" Then
                                    ItemString = ItemString & "'" & FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) & "',"
                        End If
                        
             
              Else
                            ItemString = ItemString & "'" & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & "',"
             End If
                    '    If FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) <> "" Then
                                 
                    '    Else
                    '
            '            End If
            '

          End If
          
          Next RowNum
          
 End With
 If ItemString <> "" Then
 ItemString = mId(ItemString, 1, Len(ItemString) - 1)
 Else
 ItemString = "'0'"
End If
If SystemOptions.WorkWithItemsDetails = True Then
StrSQL = "SELECT     SUM(dbo.ItemsDetails.[Count] * dbo.ItemsDetails.EffectN) AS qty, dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ClassId, dbo.ItemsDetails.UnitID, "
StrSQL = StrSQL & "                       dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.ItemId, dbo.ItemsDetails.ItemDetailedCode, dbo.TblUnites.UnitName,"
StrSQL = StrSQL & "                      dbo.TblUnites.UnitNamee"
StrSQL = StrSQL & "  FROM         dbo.ItemsDetails INNER JOIN"
StrSQL = StrSQL & "                       dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID"
StrSQL = StrSQL & " WHERE     (  dbo.Transactions.StoreID = " & val(DCboStoreName.BoundText) & ") AND (dbo.Transactions.Transaction_Date <=" & SQLDate(DTPickerAccTo.value, True) & " )"

StrSQL = StrSQL & " and  ParrtNoCode not in(" & ItemString & ")"
 
StrSQL = StrSQL & " GROUP BY dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ClassId, dbo.ItemsDetails.UnitID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.ItemId,"
StrSQL = StrSQL & "                      dbo.ItemsDetails.ItemDetailedCode , dbo.TblUnites.Unitname, dbo.TblUnites.UnitNamee"
StrSQL = StrSQL & " Having (SUM(dbo.ItemsDetails.[count] * dbo.ItemsDetails.EffectN) <> 0)"

       StrSQL = StrSQL & " order by   dbo.ItemsDetails.ItemId, Dbo.ItemsDetails.ColorID, dbo.ItemsDetails.SizeID  "
       
Else
    StrSQL = "SELECT   SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) AS SUMQTY, dbo.Transaction_Details.Item_ID AS ItemID, "
    StrSQL = StrSQL & "  dbo.Transactions.StoreID, dbo.TblStore.StoreName, 1, dbo.TblItems.ItemCode,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemName, 1, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ClassId,"
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
    StrSQL = StrSQL + " where    TblItems.ItemType=0 AND    dbo.Transactions.Transaction_Date >=" & SQLDate(DTPickerAccFrom, True) & ""
    StrSQL = StrSQL + " and dbo.Transactions.Transaction_Date <=" & SQLDate(DTPickerAccTo, True) & ""
    StrSQL = StrSQL & "   AND (dbo.Transactions.StoreID =" & val(DCboStoreName.BoundText) & ")"
StrSQL = StrSQL & " and  TblItems.ITEMID not in(" & ItemString & ")"
     
     
    StrSQL = StrSQL & "  GROUP BY dbo.Transaction_Details.Item_ID, dbo.Transactions.StoreID, dbo.TblStore.StoreName  ,"
    StrSQL = StrSQL & "  dbo.TblItems.ItemCode, dbo.TblItems.ItemName , dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize,"
    StrSQL = StrSQL & "   dbo.Transaction_Details.ClassId , dbo.TblItemsclasses.SizeName, dbo.TblItemsSizes.SizeName, dbo.TblItemsColors.ColorName"
  StrSQL = StrSQL & "  Having(SUM(dbo.Transaction_Details.Quantity * dbo.TransactionTypes.StockEffect) <> 0)"
    StrSQL = StrSQL & "  ORDER BY  ItemID"



End If

    Set RsDetails = New ADODB.Recordset
    
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText


Dim X As Integer

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
  '      Fg.Rows = RsDetails.RecordCount + 1
If RsDetails.RecordCount > 0 Then
X = MsgBox("ÌÊÃœ «’‰«› Ê€Ì— „–ﬂÊ—… »«·Ã—œ Â·  —Ìœ «÷«› Â«", vbCritical + vbYesNo)

If X = vbNo Then
Exit Function
End If


Else
Exit Function
End If


        For RowNum = 1 To RsDetails.RecordCount

            With FG
            
                                If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
                                    .rows = .rows + 1
                                End If
        
                .TextMatrix(.rows - 1, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemId").value), "", RsDetails("ItemId").value)
                .TextMatrix(.rows - 1, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemId").value), "", RsDetails("ItemId").value)
                .TextMatrix(.rows - 1, FG.ColIndex("Count")) = 0
                .TextMatrix(.rows - 1, FG.ColIndex("Serial")) = "" 'IIf(IsNull(RsDetails("ItemSerial").value), "", RsDetails("ItemSerial").value)
                .TextMatrix(.rows - 1, FG.ColIndex("HaveSerial")) = "" ' IIf(IsNull(RsDetails("HaveSerial").value), "", RsDetails("HaveSerial").value)

                 
                    FG.TextMatrix(.rows - 1, FG.ColIndex("ItemCase")) = "" '  IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
                 
                .TextMatrix(.rows - 1, FG.ColIndex("Price")) = 0
                .TextMatrix(.rows - 1, FG.ColIndex("Valu")) = 0
          FG.TextMatrix(.rows - 1, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
          If SystemOptions.WorkWithItemsDetails = True Then
            FG.TextMatrix(.rows - 1, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("SizeID")), 1, (RsDetails("SizeID").value))
                 FG.TextMatrix(.rows - 1, FG.ColIndex("ParrtNoCode")) = IIf(IsNull(RsDetails("ParrtNoCode")), "", (RsDetails("ParrtNoCode").value))
FG.TextMatrix(.rows - 1, FG.ColIndex("ItemDetailedCode")) = IIf(IsNull(RsDetails("ItemDetailedCode")), "", (RsDetails("ItemDetailedCode").value))

          Else
          FG.TextMatrix(.rows - 1, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
          End If
          
            FG.TextMatrix(.rows - 1, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
         
        Dim UnitName As String
        Dim unotfactor As Double
        
   GetDefaultItemUnit val(.TextMatrix(.rows - 1, FG.ColIndex("Code"))), LngUnitID, UnitName, unotfactor, val(FG.TextMatrix(FG.rows - 1, FG.ColIndex("Code")))
   
            FG.cell(flexcpData, .rows - 1, FG.ColIndex("UnitID")) = LngUnitID ' IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
    
    
            FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = UnitName ' IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
      
      
 

  FG.TextMatrix(.rows - 1, FG.ColIndex("AutoDetect")) = 1 ' IIf(IsNull(RsDetails("AutoDetect")), "", (RsDetails("AutoDetect").value))
  
            End With

            
            RsDetails.MoveNext

       
        Next RowNum

        FG.AutoSize 0, FG.Cols - 1, False
    End If
End Function
Function addrow(Barcode As String, Qty As Double)


If 1 = 1 Then
 '   On Error GoTo ErrTrap
    Dim StrSQL As String
'    Dim RsTemp As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngUnitID As Long
  Dim ColorID As Integer
   Dim sizeid As Integer
    Dim ClassId As Integer
    Dim ParrtNoCode As String
    Dim ItemDetailedCode As String

Dim Price As Double
  
'StrSQL = "SELECT     dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ProductionDate, dbo.ItemsDetails.ExpireDate, dbo.ItemsDetails.ColorID, "
'StrSQL = StrSQL & "                       dbo.ItemsDetails.unitid , dbo.ItemsDetails.sizeid, dbo.ItemsDetails.ClassId, dbo.ItemsDetails.ItemID, dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName"
'StrSQL = StrSQL & " FROM         dbo.ItemsDetails LEFT OUTER JOIN"
'StrSQL = StrSQL & "                       dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID"
'StrSQL = StrSQL & "  GROUP BY dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.UnitID, dbo.ItemsDetails.SizeID,"
'StrSQL = StrSQL & "                        dbo.ItemsDetails.ClassId , dbo.ItemsDetails.ProductionDate, dbo.ItemsDetails.ExpireDate, dbo.ItemsDetails.ItemID, dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName"
'

'StrSQL = StrSQL & "  HAVING        (ParrtNoCode = '" & Barcode & "')"
'
     
       
 '   Set RsTemp = New ADODB.Recordset
 '   RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText



    If Barcode <> "" Then
        RsTemp.Find "ParrtNoCode='" & Trim(Barcode) & "'", , adSearchForward, adBookmarkFirst

 


    If Not (RsTemp.EOF Or RsTemp.BOF) Then
  '      LngItemID = val(DCboItemCode.BoundText)
        LngItemID = IIf(IsNull(RsTemp("itemid").value), 0, RsTemp("ItemID").value)
  LngUnitID = IIf(IsNull(RsTemp("UnitID").value), 0, RsTemp("UnitID").value)
  ColorID = IIf(IsNull(RsTemp("ColorID").value), 0, RsTemp("ColorID").value)
  sizeid = IIf(IsNull(RsTemp("SizeID").value), 0, RsTemp("SizeID").value)
 ClassId = IIf(IsNull(RsTemp("ClassId").value), 0, RsTemp("ClassId").value)
    ParrtNoCode = IIf(IsNull(RsTemp("ParrtNoCode").value), "", RsTemp("ParrtNoCode").value)
        ItemDetailedCode = IIf(IsNull(RsTemp("ItemDetailedCode").value), "", RsTemp("ItemDetailedCode").value)
        
        If LngItemID <> 0 Then
         
    With Me.FG

        If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
            .rows = .rows + 1
        End If
                        .TextMatrix(.rows - 1, FG.ColIndex("Code")) = LngItemID
                .TextMatrix(.rows - 1, FG.ColIndex("Name")) = LngItemID
                .TextMatrix(.rows - 1, FG.ColIndex("Count")) = Qty
                .TextMatrix(.rows - 1, FG.ColIndex("Serial")) = "" ' IIf(IsNull(RsDetails("ItemSerial").value), "", RsDetails("ItemSerial").value)
                .TextMatrix(.rows - 1, FG.ColIndex("HaveSerial")) = "" ' IIf(IsNull(RsDetails("HaveSerial").value), "", RsDetails("HaveSerial").value)

               
                    FG.TextMatrix(.rows - 1, FG.ColIndex("ItemCase")) = "" ' IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
                

 

            FG.TextMatrix(.rows - 1, FG.ColIndex("ColorID")) = ColorID ' IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ItemSize")) = sizeid ' IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ClassID")) = ClassId ' IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            FG.cell(flexcpData, .rows - 1, FG.ColIndex("UnitID")) = LngUnitID ' IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(.rows - 1, FG.ColIndex("ParrtNoCode")) = ParrtNoCode
            FG.TextMatrix(.rows - 1, FG.ColIndex("ItemDetailedCode")) = ItemDetailedCode
            
            
            '    .TextMatrix(.Rows - 1, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, 0, LngUnitID)
                .TextMatrix(.rows - 1, FG.ColIndex("Price")) = 0
         .TextMatrix(.rows - 1, FG.ColIndex("Valu")) = val(.TextMatrix(.rows - 1, .ColIndex("Price"))) * val(.TextMatrix(.rows - 1, .ColIndex("Count")))

If SystemOptions.UserInterface = ArabicInterface Then
             FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsTemp("UnitName")), "", (RsTemp("UnitName").value))
Else
    FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(RsTemp("UnitNamee")), "", (RsTemp("UnitNamee").value))
End If
      
                   'AddItemInGrid .Rows - 1, DcboItemName.BoundText, val(TxtQuantity.text), val(Me.TxtPrice.text), Me.CboItemCase.ListIndex + 1, , , , , , , ColorID, sizeid, ClassId, ParrtNoCode, ItemDetailedCode

     End With
           
           
 
         'DCboItemCode_KeyDown vbKeyReturn, 0
          Me.TxtItemCodeB.text = ""
          Unload FrmItemSearch2
      Me.TxtItemCodeB.SetFocus
         
    Else
        
        
        
        

        
        
         
    End If
    
    Else
           error_string = error_string & Trim(Barcode) & "," & Qty & vbCrLf

End If
End If
'Call XPDtbBill_Change

       

End If

End Function
Function GetUnitID(Optional Name As String, Optional code As String = "") As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     UnitID"
sql = sql & " From dbo.TblUnites"
sql = sql & " WHERE     ((UnitName = N'" & Name & "') or (UnitNamee LIKE N'" & Name & "'))"
sql = sql & " and UnitID In (Select UnitId from tblItemsUnits where ItemId In (Select tblItems.itemId from tblItems where (dbo.TblItems.barcodeno = N'" & code & "') ))"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.EOF Then
    rs2.Close
    sql = " SELECT     UnitID"
    sql = sql & " From dbo.TblUnites"
    sql = sql & " WHERE     ((UnitName = N'" & Name & "') or (UnitNamee LIKE N'" & Name & "'))"
    sql = sql & " and UnitID In (Select UnitId from tblItemsUnits where ItemId In (Select tblItems.itemId from tblItems where (dbo.TblItems.FullCode = N'" & code & "') ))"
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
End If

If rs2.EOF Then

    sql = " SELECT     UnitID"
    sql = sql & " From dbo.TblUnites"
    sql = sql & " WHERE     ((UnitName LIKE N'%" & Name & "%') or (UnitNamee LIKE N'%" & Name & "%'))"
    sql = sql & " and UnitID In (Select UnitId from tblItemsUnits where ItemId In (Select tblItems.itemId from tblItems where (dbo.TblItems.barcodeno = N'" & code & "') Or (dbo.TblItems.FullCode = N'" & code & "') ))"
    rs2.Close
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
End If
If rs2.RecordCount > 0 Then
GetUnitID = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
Else
GetUnitID = 0
End If

End Function

 Public Function GetDefaultItemUnit(fullcode As String, _
                                   Optional ByRef UnitID As Long, _
                                   Optional ByRef UnitName As String, Optional ByRef UnitFactor As Double, Optional ByVal ItemID As Long = 0)
    Dim RsUnitData As New ADODB.Recordset
    
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName,TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
    Else
        StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitNamee UnitName," & "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
    End If
    StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID = TblUnites.UnitID"
    StrSQL = StrSQL + " Inner join tblItems On tblItems.ItemId =TblItemsUnits.ItemID "
    StrSQL = StrSQL + " Where 1 = 1 "
    If ItemID = 0 Then
        StrSQL = StrSQL + " and  tblItems.itemcode = N'" & Trim(fullcode) & "'"
    Else
        StrSQL = StrSQL + " and tblItems.ItemID=" & val(ItemID)
    End If
    StrSQL = StrSQL + " AND DefaultUnit=1"
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
        UnitID = IIf(IsNull(RsUnitData("UnitID").value), 0, RsUnitData("UnitID").value)
        UnitName = IIf(IsNull(RsUnitData("UnitName").value), "", RsUnitData("UnitName").value)
        UnitFactor = IIf(IsNull(RsUnitData("UnitFactor").value), 0, RsUnitData("UnitFactor").value)
    End If

    RsUnitData.Close
    Set RsUnitData = Nothing
         
End Function



Function addrow2(fullcode As String, Qty As Double, Optional UnitName As String, Optional mPrice As Double = 0, Optional ExpireDate As String, Optional sizename As String = "", Optional colorname As String = "", Optional mWidth As String = "", Optional mLength As String = "", Optional mHeight As String = "", Optional mArea As String = "")
    
    'ExpireDate, SizeName, ColorName, mWidth, mLength, mHeight, mArea
    Dim StrSQL As String
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim UnitID As Long
    Dim LngItemID As Long
    Dim LngUnitID As Long
    Dim ColorID As Integer
    Dim sizeid As Integer
    Dim ClassId As Integer
    Dim ParrtNoCode As String
    Dim ItemDetailedCode As String
  
    Dim Price As Double
    Dim mCode As String
    If chkNew.value = vbChecked Then
        'GetDefaultItemUnit Fullcode, UnitID
        UnitID = GetUnitID(UnitName, fullcode)
    Else
        UnitID = GetUnitID(UnitName, fullcode)
    End If
    If Trim(UnitName) = "" Or UnitID = 0 Then
        UnitID = 0
        GetDefaultItemUnit fullcode, UnitID
    End If
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    If sizename <> "" Then
        s = "Select sizeid,sizename from TblItemsSizes where sizename Like '%" & sizename & "%' "
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            sizeid = val(rsDummy!sizeid & "")
        End If
        rsDummy.Close
    Else
        sizeid = 1
    End If
    
    
    
    If colorname <> "" Then
        s = "Select ColorID,ColorName from TblItemsColors where ColorName Like '%" & colorname & "%' "
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            ColorID = val(rsDummy!ColorID & "")
        End If
        rsDummy.Close
    Else
        ColorID = 1
    End If
    
    
    

    
   If fullcode <> "" And UnitID <> 0 Then
     StrSQL = "  SELECT     dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, dbo.TblUnites.UnitID, dbo.TblItemsUnits.ItemID"
     StrSQL = StrSQL & "    FROM         dbo.TblItems INNER JOIN"
     StrSQL = StrSQL & "                   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID INNER JOIN"
      StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
     StrSQL = StrSQL & "  WHERE     (dbo.TblItems.barcodeno = N'" & fullcode & "') AND (dbo.TblItemsUnits.UnitID = " & UnitID & ")"
     StrSQL = StrSQL & "  GROUP BY dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, dbo.TblUnites.UnitID, dbo.TblItemsUnits.ItemID"
rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

If rs2.RecordCount = 0 Then

        StrSQL = "  SELECT     dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, dbo.TblUnites.UnitID, dbo.TblItemsUnits.ItemID"
     StrSQL = StrSQL & "    FROM         dbo.TblItems INNER JOIN"
     StrSQL = StrSQL & "                   dbo.TblItemsUnits ON dbo.TblItems.ItemID = dbo.TblItemsUnits.ItemID INNER JOIN"
      StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
     StrSQL = StrSQL & "  WHERE     (dbo.TblItems.fULLCODE = N'" & fullcode & "') AND (dbo.TblItemsUnits.UnitID = " & UnitID & ")"
     StrSQL = StrSQL & "  GROUP BY dbo.TblUnites.UnitNamee, dbo.TblUnites.UnitName, dbo.TblItems.Fullcode, dbo.TblUnites.UnitID, dbo.TblItemsUnits.ItemID"
rs2.Close
rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    
    ColorID = ColorID
End If
If rs2.RecordCount = 0 Then
    ColorID = ColorID
End If
    If rs2.RecordCount > 0 Then

         LngItemID = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
         LngUnitID = IIf(IsNull(rs2("UnitID").value), 0, rs2("UnitID").value)
         'ColorID = 1 ' IIf(IsNull(RsTemp("ColorID").value), 0, RsTemp("ColorID").value)
          'IIf(IsNull(RsTemp("SizeID").value), 0, RsTemp("SizeID").value)
         ClassId = 1 'IIf(IsNull(RsTemp("ClassId").value), 0, RsTemp("ClassId").value)
         ParrtNoCode = "" 'IIf(IsNull(RsTemp("ParrtNoCode").value), "", RsTemp("ParrtNoCode").value)
        ItemDetailedCode = "" 'IIf(IsNull(RsTemp("ItemDetailedCode").value), "", RsTemp("ItemDetailedCode").value)
        
        If LngItemID <> 0 Then
         
    With Me.FG

        If .TextMatrix(.rows - 1, .ColIndex("Code")) <> "" Then
            .rows = .rows + 1
        End If
        .TextMatrix(.rows - 1, FG.ColIndex("Code")) = LngItemID
        .TextMatrix(.rows - 1, FG.ColIndex("Name")) = LngItemID
        .TextMatrix(.rows - 1, FG.ColIndex("Count")) = Qty
        .TextMatrix(.rows - 1, FG.ColIndex("Serial")) = "" ' IIf(IsNull(RsDetails("ItemSerial").value), "", RsDetails("ItemSerial").value)
        .TextMatrix(.rows - 1, FG.ColIndex("HaveSerial")) = "" ' IIf(IsNull(RsDetails("HaveSerial").value), "", RsDetails("HaveSerial").value)
         FG.TextMatrix(.rows - 1, FG.ColIndex("ItemCase")) = "" ' IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
         FG.TextMatrix(.rows - 1, FG.ColIndex("ColorID")) = ColorID ' IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
         FG.TextMatrix(.rows - 1, FG.ColIndex("ItemSize")) = sizeid ' IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
         FG.TextMatrix(.rows - 1, FG.ColIndex("ClassID")) = ClassId ' IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
         FG.cell(flexcpData, .rows - 1, FG.ColIndex("UnitID")) = LngUnitID ' IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
         FG.TextMatrix(.rows - 1, FG.ColIndex("ParrtNoCode")) = ParrtNoCode
         FG.TextMatrix(.rows - 1, FG.ColIndex("ItemDetailedCode")) = ItemDetailedCode
         
        If mPrice = 0 Then
            If SystemOptions.CostStarting = True Then
                        Dim FirstPeriodDateInthisYear  As Date
                        getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
    
                    fromcostdate = DateAdd("d", -1, FirstPeriodDateInthisYear)
                    fromcostdate = Replace(Format$(fromcostdate, "MM/DD/yyyy"), "-", "/")
                    mPrice = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , fromcostdate, DTPickerAccTo, , LngUnitID)
              Else
                   
                    mPrice = ModItemCostPrice.GetCostItemPrice(LngItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , DTPickerAccTo, , LngUnitID)
                End If
        End If


        .TextMatrix(.rows - 1, FG.ColIndex("Price")) = mPrice
        .TextMatrix(.rows - 1, FG.ColIndex("ExpiryDate")) = ExpireDate
        
        
        FG.TextMatrix(.rows - 1, FG.ColIndex("Height")) = mHeight

        FG.TextMatrix(.rows - 1, FG.ColIndex("length")) = mLength
        FG.TextMatrix(.rows - 1, FG.ColIndex("Width")) = mWidth
            
            
        .TextMatrix(.rows - 1, FG.ColIndex("Valu")) = val(.TextMatrix(.rows - 1, .ColIndex("Price"))) * val(.TextMatrix(.rows - 1, .ColIndex("Count")))
If SystemOptions.UserInterface = ArabicInterface Then
             FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(rs2("UnitName")), "", (rs2("UnitName").value))
Else
    FG.TextMatrix(.rows - 1, FG.ColIndex("UnitID")) = IIf(IsNull(rs2("UnitNamee")), "", (rs2("UnitNamee").value))
End If

     End With
          Me.TxtItemCodeB.text = ""
     
          Unload FrmItemSearch2
      Me.TxtItemCodeB.SetFocus
         
    Else
         
    End If
    
    Else
           error_string = error_string & Trim(fullcode) & "," & Qty & "," & Name & vbCrLf

End If
End If

End Function

Private Sub XPCboGroup_Click(Area As Integer)
    If Me.TxtModFlg.text = "R" Or Me.TxtModFlg.text = "" Then Exit Sub
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
 

            Retrive2 val(Me.XPCboGroup.BoundText)

         
End Sub

Private Sub XPTxtSum_Change()
    Me.LblTotal.Caption = XPTxtSum.text
    Exit Sub
ErrTrap:
End Sub

