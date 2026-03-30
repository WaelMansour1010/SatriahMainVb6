VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRequestItems 
   Caption         =   "«·—’Ìœ «·«›  «ÕÌ"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13395
   HelpContextID   =   90
   Icon            =   "FrmRequestItems.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmRequestItems.frx":038A
   RightToLeft     =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   13395
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   7845
      Left            =   0
      TabIndex        =   7
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
      _GridInfo       =   $"FrmRequestItems.frx":0714
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   5
         Left            =   15
         TabIndex        =   11
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
            Height          =   360
            Left            =   10125
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
            Left            =   3570
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
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   315
            Index           =   63
            Left            =   8070
            TabIndex        =   67
            Top             =   120
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
            Height          =   375
            Left            =   7140
            TabIndex        =   66
            Top             =   0
            Width           =   885
         End
         Begin VB.Label lblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   10380
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   60
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·—’Ìœ"
            Height          =   255
            Index           =   3
            Left            =   11700
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   120
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   315
            Index           =   6
            Left            =   5610
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   105
            Width           =   870
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   240
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   135
            Width           =   825
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   270
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   105
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   360
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
            Height          =   480
            Index           =   4
            Left            =   2970
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   0
            Width           =   570
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   4785
         Index           =   3
         Left            =   15
         TabIndex        =   9
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
         GridRows        =   3
         GridCols        =   3
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmRequestItems.frx":078C
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin MSComctlLib.Toolbar TBr 
            Height          =   630
            Left            =   495
            TabIndex        =   26
            Top             =   4395
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   1111
            ButtonWidth     =   609
            ButtonHeight    =   1005
            Appearance      =   1
            _Version        =   393216
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   690
            Index           =   4
            Left            =   30
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   30
            Width           =   13305
            _cx             =   23469
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
               Left            =   675
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   300
               Width           =   1635
            End
            Begin VB.TextBox TxtSerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   4050
               MaxLength       =   20
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   300
               Width           =   2055
            End
            Begin VB.TextBox TxtQuantity 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   2430
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   300
               Width           =   1530
            End
            Begin VB.ComboBox CboItemCase 
               Height          =   315
               Left            =   6150
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   300
               Width           =   1905
            End
            Begin MSDataListLib.DataCombo DCboItemsName 
               Height          =   315
               Left            =   8055
               TabIndex        =   1
               Top             =   270
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
               TabIndex        =   0
               Top             =   300
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdAdd 
               Height          =   420
               Left            =   30
               TabIndex        =   6
               Top             =   210
               Width           =   420
               _ExtentX        =   741
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
               ButtonImage     =   "FrmRequestItems.frx":07EE
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
               Left            =   855
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   30
               Width           =   1455
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   270
               Index           =   27
               Left            =   2655
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   30
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ì—Ì«·"
               Height          =   390
               Index           =   28
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   30
               Width           =   1950
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·… «·’‰›"
               Height          =   270
               Index           =   29
               Left            =   6285
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   30
               Width           =   1770
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈”„ «·’‰›"
               Height          =   270
               Index           =   30
               Left            =   8280
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   30
               Width           =   2370
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·’‰›"
               Height          =   270
               Index           =   31
               Left            =   10875
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   30
               Width           =   2415
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   3645
            Left            =   30
            TabIndex        =   65
            Top             =   735
            Width           =   13305
            _cx             =   23469
            _cy             =   6429
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
            FormatString    =   $"FrmRequestItems.frx":0B88
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
            Top             =   4395
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
         Begin VB.TextBox txtopening_balance_voucher_id 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2910
            RightToLeft     =   -1  'True
            TabIndex        =   60
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
            Left            =   -30
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   0
            Width           =   7920
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
            Begin MSDataListLib.DataCombo DCboStoreName 
               Height          =   315
               Left            =   1680
               TabIndex        =   63
               Top             =   870
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Œ“‰"
               Height          =   375
               Index           =   2
               Left            =   4485
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   840
               Width           =   945
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
            Left            =   7065
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   90
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   60
            Width           =   2160
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
            Left            =   9360
            TabIndex        =   30
            Top             =   480
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   44433411
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   9360
            TabIndex        =   61
            Top             =   960
            Width           =   2145
            _ExtentX        =   3784
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
            Left            =   12495
            TabIndex        =   62
            Top             =   960
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   375
            Index           =   0
            Left            =   11520
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   465
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”·”·"
            Height          =   375
            Index           =   1
            Left            =   11520
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   75
            Width           =   1680
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   615
         Left            =   15
         TabIndex        =   45
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
         Caption         =   "«·—’Ìœ «·«›  «ÕÌ"
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
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1725
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
            ButtonImage     =   "FrmRequestItems.frx":0E18
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
            ButtonImage     =   "FrmRequestItems.frx":11B2
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
            ButtonImage     =   "FrmRequestItems.frx":154C
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
            ButtonImage     =   "FrmRequestItems.frx":18E6
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
         Height          =   540
         Index           =   0
         Left            =   0
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   7305
         Width           =   13395
         _cx             =   23627
         _cy             =   952
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
            Left            =   12000
            TabIndex        =   51
            Top             =   105
            Width           =   1275
            _ExtentX        =   2249
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
            Left            =   10545
            TabIndex        =   52
            Top             =   105
            Width           =   1155
            _ExtentX        =   2037
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
            Left            =   9000
            TabIndex        =   53
            Top             =   120
            Width           =   1425
            _ExtentX        =   2514
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
            Left            =   7620
            TabIndex        =   54
            Top             =   105
            Width           =   1215
            _ExtentX        =   2143
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
            Left            =   5790
            TabIndex        =   55
            Top             =   105
            Width           =   1680
            _ExtentX        =   2963
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
            Left            =   4545
            TabIndex        =   56
            Top             =   105
            Width           =   1170
            _ExtentX        =   2064
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
            TabIndex        =   57
            Top             =   105
            Width           =   1365
            _ExtentX        =   2408
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
            Left            =   2910
            TabIndex        =   58
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
            Left            =   1620
            TabIndex        =   59
            Top             =   105
            Width           =   1155
            _ExtentX        =   2037
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
Attribute VB_Name = "FrmRequestItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim BalanceReport As ClsOpeningBalanceReport
Dim cSearchDcbo As clsDCboSearch
Dim NewGrid As New ClsGrid
 Dim FirstPeriodDateInthisYear  As Date
 



Private Sub CmdHelp_Click()
SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 8
        FrmItemSearch.Show vbModal
End If
End Sub

Private Sub DCboStoreName_Change()
WriteDev
End Sub

Private Sub DCboStoreName_Click(Area As Integer)
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
fill_combo DcBranch, My_SQL
 
 If SystemOptions.usertype <> UserAdminAll Then
 
 Me.DcBranch.Enabled = False
End If


Set Cmd(0).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("New").Picture
Set Cmd(1).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Edit").Picture
Set Cmd(2).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("save").Picture
Set Cmd(3).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Undo").Picture
Set Cmd(4).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Del").Picture
Set Cmd(5).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Search").Picture
Set Cmd(6).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Exit").Picture
Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Print").Picture
Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture

Resize_Form Me, TransactionSize
If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    changelang
End If


FG.WallPaper = BGround.Picture
AddTip
SetDtpickerDate XPDtbBill
NewGrid.GridTrans = OpeningBalance
Set NewGrid.TxtFillData = TxtFillData
Set NewGrid.Grid = FG
Set NewGrid.TxtModFlag = TxtModFlg
Set NewGrid.StoreName = Me.DCboStoreName
Set NewGrid.LblItemsCount = Me.LblItemsCount
' ⁄»∆… »Ì«‰«  «·√’‰«›
Set NewGrid.DcboItemName = DCboItemsName
Set NewGrid.DCboItemCode = DCboItemsCode
Set NewGrid.CboItemCase = CboItemCase
Set NewGrid.CmdAddData = CmdAdd
Set NewGrid.TxtSerial = TxtSerial
Set NewGrid.TxtQuantity = TxtQuantity
Set NewGrid.TxtPrice = TxtPrice
Set NewGrid.GrdTBar = Me.TBr
' Set NewGrid.LblTotalQty = Me.LblTotalQty
Set NewGrid.Txttotal = Me.XPTxtSum
Set NewGrid.TxtInvID = Me.XPTxtBillID
Set NewGrid.LblTotalQty = Me.LblTotalQty
NewGrid.FillGrid
Set Dcombos = New ClsDataCombos
Dcombos.GetUsers Me.DCboUserName
Dcombos.GetAccountingCodes Me.DcboDebitSide
Dcombos.GetAccountingCodes Me.DcboCreditSide
Dcombos.GetStores Me.DCboStoreName
Set cSearchDcbo = New clsDCboSearch
Set cSearchDcbo.Client = Me.DCboStoreName

StrSQL = "Select * From Transactions where Transaction_Type=3"
Set rs = New ADODB.Recordset
rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

XPBtnMove_Click 2
TxtModFlg.text = "R"
If OPEN_NEW_SCREEN = True Then
Cmd_Click (0)
End If


Exit Sub
ErrTrap:
Msg = Err.description
Msg = Msg & Chr(13) & Err.Number
Msg = Msg & Chr(13) & Err.Source
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

getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

 Me.XPDtbBill.value = FirstPeriodDateInthisYear
 
Dim intDef As Integer
Select Case Index
    Case 0
        If DoPremis(Do_New, Me.name, True) = False Then
            Exit Sub
        End If
        clear_all Me
        TxtModFlg.text = "N"
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Me.TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=3"))
        txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
        WriteDev
        GridDefaultValue FG.Rows - 1
        Me.DCboUserName.BoundText = user_id
        intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
        DCboStoreName.BoundText = intDef
        FG.SetFocus
        FG.Rows = 2
        FG.Col = FG.ColIndex("Code")
        FG.Row = FG.Rows - 1
        Me.DcBranch.BoundText = branch_id
    Case 1
        If DoPremis(Do_Edit, Me.name, True) = False Then
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
        If DoPremis(Do_Delete, Me.name, True) = False Then
            Exit Sub
        End If
        Del_TransAction
    Case 7
        If DoPremis(Do_Print, Me.name, True) = False Then
            Exit Sub
        End If
                AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)
        If AskOption = False Then
            FrmPrintOptions.Show vbModal
        End If
     '   If BolPrint = False Then
     '       Exit Sub
     '   End If


        printing
    Case 5
        If DoPremis(Do_Search, Me.name, True) = False Then
            Exit Sub
        End If
        FrmBalanceSearch.Show vbModal
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
        Me.DCboStoreName.Locked = True
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
        Me.DCboStoreName.Locked = False
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
        Me.DCboStoreName.Locked = False
        FG.Editable = flexEDKbdMouse
        Ele(4).Enabled = True
End Select
Exit Sub
ErrTrap:
End Sub
Private Sub Del_TransAction()
On Error GoTo ErrTrap
If XPTxtBillID.text <> "" Then
    Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
    Msg = Msg + (XPTxtBillID.text) & Chr(13)
    Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
        If AvailableDeal = True Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                    
           StrSqlDel = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & Val(txtopening_balance_voucher_id.text)
             Cn.Execute StrSqlDel, , adExecuteNoRecords
       
     '  Update_opening_balance_screen_accounts
'       MsgBox " „ «·Õ–›"
       
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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + _
            vbExclamation, App.Title
    rs.CancelUpdate
End If
End Sub
Private Sub AddTip()
Dim Wrap As String
On Error GoTo ErrTrap
Wrap = Chr(13) + Chr(10)
Set TTP = New clstooltip
With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(0), _
    "ÃœÌœ ..." & Wrap & _
    "·«÷«›… »Ì«‰«   ÃœÌœ…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(7), _
    "ÿ»«⁄… ..." & Wrap & _
    "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & _
    " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
End With
With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(1), _
    " ⁄œÌ· ..." & Wrap & _
    "· ⁄œÌ· Â–Â «·»Ì«‰« " & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(2), _
    "Õ›Ÿ ..." & Wrap & _
    "·Õ›Ÿ Â–Â «·»Ì«‰« " & Wrap & _
     "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(3), _
    " —«Ã⁄ ..." & Wrap & _
    "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & _
     "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
 With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(4), _
    "Õ–› ..." & Wrap & _
    "·Õ–› Â–Â «·»Ì«‰« " & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(5), _
    "»ÕÀ ..." & Wrap & _
    "···»ÕÀ ⁄‰ ⁄„·Ì… " & Wrap & _
    "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl Cmd(6), _
    "Œ—ÊÃ ..." & Wrap & _
    "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With

With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(1), _
    "«·√Ê· ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(0), _
    "«·”«»ﬁ ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(3), _
    "«· «·Ì ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl XPBtnMove(2), _
    "«·√ŒÌ— ..." & Wrap & _
    "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & _
    " ›ﬁÿ ≈÷€ÿ Â‰«", True
End With
With TTP
   .Create Me.hWnd, "«·—’Ìœ «·«›  «ÕÌ", 1, 15204351, -2147483630
   .MaxWidth = 4000
   .VisibleTime = 9000
   .DelayTime = 600
   .AddControl CmdHelp, _
    "„”«⁄œ… ..." & Wrap & _
    "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & _
    "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & _
    "≈÷€ÿ Â‰«" & Wrap, True
End With
Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
On Error GoTo ErrTrap
Select Case TxtModFlg.text
    Case "N"
         Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
        Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
        If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)
            
           End If
    Case "E"
    Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
        Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
        If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
         rs.Find "Transaction_ID='" & Val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst
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

    If Trim(DcBranch.BoundText) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify Departement"
            Else
                Msg = "ÌÃ»  ÕœÌœ «”„    «·›—⁄"
            End If
   MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    DcBranch.SetFocus
    SendKeys "{F4}"
    Screen.MousePointer = vbDefault
    Exit Sub
    End If
    
If Me.TxtModFlg.text <> "R" Then
    If DCboStoreName.BoundText = "" Then
        Msg = "ÌÃ» «Œ Ì«— «”„ «·„Œ“‰"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboStoreName.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If NewGrid.IsReaptedSerials = True Then
        Msg = "ÌÊÃœ  ﬂ—«— ›Ï √—ﬁ«„ «·”Ì—Ì«· «·„œŒ·… "
        Msg = Msg & Chr(13) & "»—Ã«¡ «· «ﬂœ „‰ «·√—ﬁ«„ «·„œŒ·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    If TxtModFlg.text = "N" Then
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Me.TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=3"))
        rs.AddNew
        rs("Transaction_ID").value = Val(XPTxtBillID.text)
    End If
    RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs("BranchId").value = IIf(Me.DcBranch.BoundText = "", 0, Val(DcBranch.BoundText))
    rs("opening_balance_voucher_id").value = Val(txtopening_balance_voucher_id.text)
    rs("Transaction_Serial").value = Me.TxtTransSerial.text
    rs("Transaction_Date").value = XPDtbBill.value
    rs("Transaction_Type").value = 3
    rs("UserID").value = user_id
    rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, DCboStoreName.BoundText)
    rs.update
    If Me.TxtModFlg.text = "E" Then
       StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & Val(rs("Transaction_ID").value)
       Cn.Execute StrSqlDel, , adExecuteNoRecords
      ' StrSqlDel = "delete From NOTES where Transaction_ID=" & Val(rs("Transaction_ID").value)
      ' Cn.Execute StrSqlDel, , adExecuteNoRecords
    
           StrSqlDel = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & Val(txtopening_balance_voucher_id.text)
       Cn.Execute StrSqlDel, , adExecuteNoRecords
       
    
    End If
    For RowNum = 1 To FG.Rows - 1
        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            RSTransDetails.AddNew
            RSTransDetails("Transaction_ID").value = XPTxtBillID.text
            RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
            RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
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
            RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
            RSTransDetails("Price").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            
            RSTransDetails("ColorID").value = _
            IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, Val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            
             RSTransDetails("ItemSize").value = _
            IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
              RSTransDetails("ClassId").value = _
            IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, Val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
            
           RSTransDetails("BranchId").value = IIf(Me.DcBranch.BoundText = "", 0, Val(DcBranch.BoundText))
            ' IIf((FG.TextMatrix(RowNum, FG.ColIndex("BranchId")) = ""), 1, Val(FG.TextMatrix(RowNum, FG.ColIndex("BranchId"))))
               
           ' RSTransDetails("ItemSize").value = _
            IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), "", Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
            
             RSTransDetails("UnitID").value = _
         IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
       RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
 

 Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double

        
            LngCurItemID = Val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            LngUnitID = Val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            DblQty = Val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rs.BOF Or rs.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
            End If

         RSTransDetails("Price").value = Val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                              
           RSTransDetails("showprice").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))

            RSTransDetails("OpeningBurcahseQty").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseQty")) = "", Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseQty"))))
            RSTransDetails("OpeningBurcahseValue").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseValue")) = "", Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseValue"))))
            RSTransDetails("OpeningSalesQty").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesQty")) = "", Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesQty"))))
            RSTransDetails("OpeningSalesValue").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesValue")) = "", Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesValue"))))
            RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
            RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
            RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
            
            RSTransDetails.update
        End If
    Next RowNum
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open "NOTES1", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    
   Dim NoteSerial As String
   Dim NoteSerial1 As String

NoteSerial = year(XPDtbBill.value) & 1
NoteSerial1 = NoteSerial
   
   
'    Dim NoteSerial As String
'    Dim noteserial1 As String
'
'    NoteSerial = ""
'
'                  If NoteSerial = "" Then
'                       If Notes_coding(Val(my_branch), XPDtbBill.value) = "error" Then
'                       MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
'                       Else
'
'                       If Notes_coding(Val(my_branch), XPDtbBill.value) = "" Then
'                       MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
'                       Else
'                       NoteSerial = Notes_coding(Val(my_branch), XPDtbBill.value)
'                       End If
'                       End If
'                End If
'
'                If noteserial1 = "" Then
'                   If Voucher_coding(Val(my_branch), XPDtbBill.value, 3, 1000) = "error" Then
'                   MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ﬁÌœ «›  «ÕÌ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
'                   Else
'
'                   If Voucher_coding(Val(my_branch), XPDtbBill.value, 3, 1000) = "" Then
'                   MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
'                   Else
'                   noteserial1 = Voucher_coding(Val(my_branch), XPDtbBill.value, 3, 1000)
'                   End If
'                   End If
'                End If
                
    
'    RsNotes.AddNew
'        'LngNoteID = new_id("NOTES", "NoteID", "")
'        'RsNotes("NoteID").value = LngNoteID
'        RsNotes("NoteID").value = 1
'        RsNotes("NoteDate").value = XPDtbBill.value
'        RsNotes("NoteType").value = 101
'        RsNotes("NoteSerial").value = NoteSerial ' new_id("NOTES", "NoteSerial", "", True, "NOTETYPE=100")
'          RsNotes("NoteSerial1").value = noteserial1
'
'        RsNotes("Note_Value").value = Val(lbltotal.Caption)
'        RsNotes("Transaction_ID").value = Val(Me.XPTxtBillID.text)
'    RsNotes.update
'«·œ«‘‰
    LngNoteID = 1
   ' LngDev = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    LngDev = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "")
  If SystemOptions.UserInterface = ArabicInterface Then
    StrTemp = "—’Ìœ ≈›  «ÕÏ ··„Œ«“‰  —ﬁ„ " & Trim(Me.TxtTransSerial.text) & " ·”‰Â " & year(XPDtbBill.value)
 Else
 StrTemp = "   Opening Balance No:  " & Trim(Me.TxtTransSerial.text) & " Year " & year(XPDtbBill.value)
 End If
    If ModAccounts.AddNewDev(LngDev, 1, Me.DcboCreditSide.BoundText, Val(Me.LblTotal.Caption), 1, _
        StrTemp, LngNoteID, , , CInt(SystemOptions.SysCurrentAccountIntervalID), Me.XPDtbBill.value, , , , , , , , , , , , , , True, Val(txtopening_balance_voucher_id.text), , , , Val(DcBranch.BoundText)) = False Then
        GoTo ErrTrap
    Else
  '  update_account_opening_balance Me.DcboCreditSide.BoundText
    End If
    
    
    
'    If ModAccounts.AddNewDev(LngDev, 2, Me.DcboCreditSide.BoundText, Val(Me.lbltotal.Caption), 1, _
'        StrTemp, LngNoteID, , , CInt(SystemOptions.SysCurrentAccountIntervalID), Me.XPDtbBill.value, , , , , , , , , , , , , , True, Val(txtopening_balance_voucher_id.text)) = False Then
'        GoTo ErrTrap
'    End If
    
Dim LngDevNO  As Integer
Dim StrTempAccountCode As String
Dim StrTempDes As String
Dim SngTemp  As Variant

Dim Account_Code_dynamic As String
Dim i As Integer

LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "")
LngDevNO = 0
'«·ÿ—› «·„œÌ‰
 SngTemp = (Me.LblTotal.Caption)
If SngTemp > 0 Then
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
            If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "—’Ìœ «›  «ÕÌ   ··„Œ«“‰ ·⁄«„" & year(XPDtbBill.value)
            Else
            StrTempDes = "Opening Balance Year" & year(XPDtbBill.value)
            End If
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDev, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, LngNoteID, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , True, Me.txtopening_balance_voucher_id, , , , Val(DcBranch.BoundText)) = False Then
                GoTo ErrTrap
            Else
      '      update_account_opening_balance StrTempAccountCode
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
            StrTempDes = "—’Ìœ «›  «ÕÌ ··„Œ«“‰   ·⁄«„ " & year(XPDtbBill.value)
            Else
            StrTempDes = "Opening Balance For Inventory Year" & year(XPDtbBill.value)
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDev, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, LngNoteID, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , True, Me.txtopening_balance_voucher_id, , , , Val(DcBranch.BoundText)) = False Then
                GoTo ErrTrap
           Else
      '     update_account_opening_balance StrTempAccountCode
            
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
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
             If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "—’Ìœ «›  «ÕÌ  ··„Œ«“‰  ·⁄«„  " & year(XPDtbBill.value)
            Else
            StrTempDes = "Opening Balance Year" & year(XPDtbBill.value)
            End If
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDev, LngDevNO, groupAccount, line_value, 0, StrTempDes, LngNoteID, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , True, Me.txtopening_balance_voucher_id, , , , Val(DcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                            Else
                        '    update_account_opening_balance groupAccount
                        End If
    
                    End If

                Next i

            End With

        End If
    End If
    Cn.CommitTrans
    BegineTrans = False
    Me.LblDevID.Caption = LngDev
    Me.LblAccountInterval.Caption = SystemOptions.SysCurrentAccountIntervalID
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
      
       
       
      
    Select Case Me.TxtModFlg.text
        Case "N"
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
        Else
             Msg = " Successfully Saved " & Chr(13)
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
    If SystemOptions.SysMainStockCostMethod = ModernWeightAverage Or SystemOptions.SysMainStockCostMethod = LastPurPriceType Then
        '›Ï Õ«·… «‰  ﬂÊ‰ ÿ—Ìﬁ… Õ”«» „ Ê”ÿ «· ﬂ·›…
        'ÂÊ
        'ModernWeightAverage
        '·«»œ «‰ ÌﬁÊ„ «·»—‰«„Ã » ⁄œÌ· ﬁÌ„… „ Ê”ÿ «· ﬂ·›… ··√’‰«›
        '«·„ÊÃÊœ… ›Ï «·›« Ê—…
         UpdateTransCost Val(Me.XPTxtBillID.text)
    End If
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
    Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
    Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub
End If

Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  "
Msg = Msg & Chr(13) & "" & Err.description
Msg = Msg & Chr(13) & "" & Err.Number
Msg = Msg & Chr(13) & "" & Err.Source
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
Screen.MousePointer = vbArrowHourglass
DcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", Val(rs("BranchId").value))
XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", rs("opening_balance_voucher_id").value)
Me.TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", rs("Transaction_Date").value)
DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
FG.Clear flexClearScrollable, flexClearEverything
FG.Rows = 2
StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & _
"ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
StrSQL = StrSQL + " where Transaction_ID=" & Val(rs("Transaction_ID").value)

StrSQL = StrSQL + "order by id"

'StrSql = "select * From Transaction_Details where Transaction_ID=" & Val(Rs("Transaction_ID").Value)
RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Not (RsDetails.EOF Or RsDetails.BOF) Then
    FG.Rows = RsDetails.RecordCount + 1
    For RowNum = 1 To RsDetails.RecordCount
        With FG
            .TextMatrix(RowNum, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID").value), "", RsDetails("Item_ID").value)
            .TextMatrix(RowNum, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID").value), "", RsDetails("Item_ID").value)
            .TextMatrix(RowNum, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty").value), "", RsDetails("showqty").value)
            .TextMatrix(RowNum, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial").value), "", RsDetails("ItemSerial").value)
            .TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = IIf(IsNull(RsDetails("HaveSerial").value), "", RsDetails("HaveSerial").value)
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Else
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
            End If
            .TextMatrix(RowNum, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice").value), "", RsDetails("showPrice").value)
            .TextMatrix(RowNum, FG.ColIndex("Valu")) = Val(.TextMatrix(RowNum, .ColIndex("Price"))) * Val(.TextMatrix(RowNum, .ColIndex("Count")))
        End With
        FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
        FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
        FG.TextMatrix(RowNum, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
        FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
'        Fg.TextMatrix(RowNum, Fg.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
        If SystemOptions.UserInterface = ArabicInterface Then
        FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
       Else
       FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitNamee")), "", (RsDetails("UnitNamee").value))
       End If
       
         FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseQty")) = IIf(IsNull(RsDetails("OpeningBurcahseQty").value), "", RsDetails("OpeningBurcahseQty").value)
          FG.TextMatrix(RowNum, FG.ColIndex("OpeningBurcahseValue")) = IIf(IsNull(RsDetails("OpeningBurcahseValue").value), "", RsDetails("OpeningBurcahseValue").value)
           FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesQty")) = IIf(IsNull(RsDetails("OpeningSalesQty").value), "", RsDetails("OpeningSalesQty").value)
            FG.TextMatrix(RowNum, FG.ColIndex("OpeningSalesValue")) = IIf(IsNull(RsDetails("OpeningSalesValue").value), "", RsDetails("OpeningSalesValue").value)
            FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
            FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", RsDetails("FoxyNo").value)
    
        RsDetails.MoveNext
        If FG.Rows > 10 Then
            If RowNum = 8 Then FG.Refresh
        End If
    Next RowNum
    FG.AutoSize 0, FG.Cols - 1, False
End If
Me.XPTxtSum.text = FG.Aggregate(flexSTSum, FG.FixedRows, _
    FG.ColIndex("Valu"), FG.Rows - 1, FG.ColIndex("Valu"))
Me.LblTotalQty = FG.Aggregate(flexSTSum, FG.FixedRows, _
    FG.ColIndex("Count"), FG.Rows - 1, FG.ColIndex("Count"))
    
    
StrSQL = "Select * From NOTES Where Transaction_ID=" & Val(Me.XPTxtBillID.text)
Set RsNotes = New ADODB.Recordset
RsNotes.Open StrSQL, Cn
If Not (RsNotes.BOF Or RsNotes.EOF) Then
    LngNoteID = RsNotes("NoteID").value
    StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & LngNoteID & ""
    StrSQL = StrSQL + " Order BY DEV_ID_Line_No"
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RsDev.BOF Or RsDev.EOF) Then
        Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
        Me.LblAccountInterval.Caption = RsDev("Account_Interval_ID").value
        RsDev.MoveFirst
        For i = 1 To RsDev.RecordCount
            If RsDev("Credit_Or_Debit").value = 0 Then
                Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
            ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
            End If
            RsDev.MoveNext
        Next i
    End If
    
End If
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

Dim ShowType As Boolean
ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)
If ShowType = True Then
    If Not XPTxtBillID.text Then
        Set BalanceReport = New ClsOpeningBalanceReport
        BalanceReport.ShowOpeningBalanceData XPTxtBillID.text
    End If
Else
    If Not XPTxtBillID.text Then
        Set BalanceReport = New ClsOpeningBalanceReport
BalanceReport.ShowOpeningBalanceData XPTxtBillID.text, True 'Short View
    End If
End If
 




'If XPTxtBillID.text <> "" Then
'    Set BalanceReport = New ClsOpeningBalanceReport
'    BalanceReport.ShowOpeningBalanceData XPTxtBillID.text
'End If
'Exit Sub
ErrTrap:
End Sub

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
                        .Show vbModal
                    End With
                    AvailableDeal = False
                    Exit Function
'                End If
                RsTemp.Close
            Else
                LngItemID = Val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                Set RsTemp = New ADODB.Recordset
                Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, Val(Me.XPTxtBillID.text))
                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    If Val(RsTemp("totalqty").value) < Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) Then
                        With FrmAlarm
                            .DealingForm = OpeningBalance
                            .Show vbModal
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim IntResult As String
Dim StrMSG As String
On Error GoTo ErrTrap
If Me.TxtModFlg.text <> "R" Then
Select Case Me.TxtModFlg.text
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub changelang()
Dim XPic As IPictureDisp
Set XPic = Me.XPBtnMove(1).ButtonImage
Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
Set Me.XPBtnMove(2).ButtonImage = XPic

Set XPic = Me.XPBtnMove(0).ButtonImage
Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
Set Me.XPBtnMove(3).ButtonImage = XPic


 

Me.Caption = "Opening Balance"
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
.TextMatrix(0, .ColIndex("NewItem")) = "NewItem"
End With
   
   
'NewItem

End Sub

Private Sub XPTxtSum_Change()
Me.LblTotal.Caption = XPTxtSum.text
Exit Sub
ErrTrap:
End Sub
