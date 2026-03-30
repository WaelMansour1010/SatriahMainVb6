VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmPayments1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " „ÊÌ· Ê«” ⁄«÷… «·Œ“‰ Ê «·⁄Âœ"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8130
   HelpContextID   =   390
   Icon            =   "FrmPayments1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   4695
      Left            =   0
      TabIndex        =   85
      Top             =   2160
      Width           =   8175
      _cx             =   14420
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   14871017
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "«·»Ì«‰«  «·«”«”Ì…|Õ«·… «·«⁄ „«œ"
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   4320
         Left            =   45
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   45
         Width           =   8085
         _cx             =   14261
         _cy             =   7620
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
         Begin VB.TextBox XPTxtVal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3720
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox XPMTxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   645
            Left            =   60
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   110
            Top             =   1185
            Width           =   2715
         End
         Begin VB.ComboBox CboPaymentType 
            Height          =   315
            Left            =   3720
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   825
            Width           =   2655
         End
         Begin VB.Frame FraNote 
            BackColor       =   &H00E2E9E9&
            Height          =   1965
            Left            =   3750
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   1215
            Width           =   4155
            Begin VB.TextBox TxtChequeNumber 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   810
               Width           =   2685
            End
            Begin VB.TextBox txtperson 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   1560
               Width           =   2685
            End
            Begin MSComCtl2.DTPicker DtpChequeDueDate 
               Height          =   315
               Left            =   30
               TabIndex        =   101
               Top             =   1140
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               _Version        =   393216
               Format          =   137953281
               CurrentDate     =   39614
            End
            Begin MSDataListLib.DataCombo DcboBankName 
               Height          =   315
               Left            =   30
               TabIndex        =   102
               Top             =   480
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   315
               Left            =   30
               TabIndex        =   103
               Top             =   150
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
               Height          =   285
               Index           =   17
               Left            =   2820
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   1140
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·‘Ìﬂ"
               Height          =   285
               Index           =   16
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   810
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·»‰ﬂ"
               Height          =   285
               Index           =   15
               Left            =   2790
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   510
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Œ“‰…"
               Height          =   285
               Index           =   9
               Left            =   2790
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„” ›Ìœ"
               Height          =   285
               Index           =   34
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   1560
               Width           =   1215
            End
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
            Height          =   885
            Index           =   1
            Left            =   1410
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   3405
            Width           =   6495
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   200
               Width           =   1785
            End
            Begin MSDataListLib.DataCombo DcboDebitSide 
               Height          =   315
               Left            =   90
               TabIndex        =   90
               Top             =   180
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCreditSide 
               Height          =   315
               Left            =   90
               TabIndex        =   91
               Top             =   510
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   11
               Left            =   3870
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   510
               Width           =   1485
            End
            Begin VB.Label LblDevID 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Left            =   3870
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   210
               Width           =   1485
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·› —… :"
               Height          =   315
               Index           =   29
               Left            =   5370
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   540
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ﬁÌœ:"
               Height          =   315
               Index           =   30
               Left            =   5370
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   210
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› œ«∆‰"
               Height          =   285
               Index           =   31
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   510
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› „œÌ‰"
               Height          =   285
               Index           =   32
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   180
               Width           =   885
            End
         End
         Begin VB.TextBox txt_general_des 
            Alignment       =   1  'Right Justify
            Height          =   1365
            Left            =   60
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   87
            Top             =   2085
            Width           =   2715
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   0
            TabIndex        =   112
            Top             =   15
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseAniLabel.ISAniLabel LblLink 
            Height          =   315
            Left            =   0
            TabIndex        =   113
            Top             =   360
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   556
            ActiveUnderline =   -1  'True
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   4210688
            MousePointer    =   99
            MouseIcon       =   "FrmPayments1.frx":038A
            BackColor       =   14871017
            Alignment       =   1
            Caption         =   ""
            ColorHover      =   16711680
            RightToLeft     =   -1  'True
            ImageCount      =   0
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   285
            Index           =   5
            Left            =   2730
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   1215
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„"
            Height          =   285
            Index           =   3
            Left            =   6420
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   30
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·„œ›Ê⁄« "
            Height          =   285
            Index           =   2
            Left            =   6420
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   375
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—’Ìœ «·Õ«·Ï:"
            Height          =   285
            Index           =   13
            Left            =   2460
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   375
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   18
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   735
            Width           =   3555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   315
            Index           =   14
            Left            =   6420
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   825
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—Õ «·⁄«„"
            Height          =   285
            Index           =   36
            Left            =   2580
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   2085
            Width           =   1155
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   4320
         Left            =   8820
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   45
         Width           =   8085
         _cx             =   14261
         _cy             =   7620
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
         Begin VSFlex8UCtl.VSFlexGrid GRID2 
            Height          =   3615
            Left            =   0
            TabIndex        =   122
            Tag             =   "1"
            Top             =   120
            Width           =   8055
            _cx             =   14208
            _cy             =   6376
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
            FormatString    =   $"FrmPayments1.frx":04EC
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
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   3840
            Width           =   3375
         End
         Begin VB.Label Label1100 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
            Height          =   255
            Left            =   11040
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   4560
            Width           =   3375
         End
      End
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4890
      RightToLeft     =   -1  'True
      TabIndex        =   82
      Top             =   600
      Width           =   1545
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   79
      Top             =   720
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Frame Frame2 
      Caption         =   "›Ì Õ«·… «·„ÊŸ›"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   1680
      Visible         =   0   'False
      Width           =   4215
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         Caption         =   "«ÃÊ— „” Õﬁ…"
         Height          =   195
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         Caption         =   "”·›…"
         Height          =   195
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAdv_payment_value 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   9060
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   2955
      Width           =   2685
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ŒÌ«—« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   480
      Width           =   3735
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "œ›⁄Â „ﬁœ„Â"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "FIFO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " ÕœÌœ ›Ê« Ì—"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   720
         Width           =   2055
      End
      Begin ALLButtonS.ALLButton ALLButton3 
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   " ÕœÌœ"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmPayments1.frx":062F
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
   Begin VB.Frame FraInfo 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„«   Â„ﬂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2265
      Left            =   9270
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3810
      Width           =   3705
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   0
         Left            =   1830
         TabIndex        =   35
         Top             =   780
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments1.frx":064B
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   780
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments1.frx":07AD
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   2
         Left            =   1830
         TabIndex        =   37
         Top             =   1350
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments1.frx":090F
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   38
         Top             =   1350
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments1.frx":0A71
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   4
         Left            =   1830
         TabIndex        =   39
         Top             =   1920
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments1.frx":0BD3
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   40
         Top             =   1920
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments1.frx":0D35
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   41
         Top             =   540
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments1.frx":0E97
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   42
         Top             =   1110
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments1.frx":0FF9
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin ImpulseAniLabel.ISAniLabel LblLinkInfo 
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   397
         ActiveUnderline =   -1  'True
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   4210688
         MousePointer    =   99
         MouseIcon       =   "FrmPayments1.frx":115B
         BackColor       =   14871017
         Alignment       =   1
         Caption         =   ""
         ColorHover      =   16711680
         RightToLeft     =   -1  'True
         ImageCount      =   0
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„œ›Ê⁄«  ›Ï «·≈”»Ê⁄ «·Õ«·Ï:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   19
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1110
         Width           =   2235
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„œ›Ê⁄«  ›Ï «·‘Â— «·Õ«·Ï :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   20
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1680
         Width           =   2235
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ﬁœÌ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   21
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1350
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·≈”»Ê⁄ «·Õ«·Ï"
         Height          =   255
         Index           =   22
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈Ã„«·Ï „œ›Ê⁄«  «·ÌÊ„:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   23
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   540
         Width           =   2235
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Ìﬂ« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   24
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1350
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ﬁœÌ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   25
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Ìﬂ« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   26
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰ﬁœÌ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   27
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   780
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Ìﬂ« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   28
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   780
         Width           =   675
      End
   End
   Begin VB.CheckBox ChkTrans 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰ Õ”«» ›« Ê—…"
      Height          =   225
      Left            =   9690
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   390
      Width           =   1575
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   975
      Index           =   0
      Left            =   8070
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1110
      Width           =   3675
      Begin VB.TextBox TxtTransID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TxtTransSerial 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1110
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   540
         Width           =   1005
      End
      Begin VB.ComboBox CboTrans 
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   210
         Width           =   1995
      End
      Begin ImpulseButton.ISButton CmdSearchTrans 
         Height          =   345
         Left            =   600
         TabIndex        =   19
         Top             =   540
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonPositionImage=   1
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmPayments1.frx":12BD
      End
      Begin ImpulseButton.ISButton CmdOpenTrans 
         Height          =   345
         Left            =   90
         TabIndex        =   21
         Top             =   540
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonPositionImage=   1
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmPayments1.frx":1657
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«œŒ· —ﬁ„ «·›« Ê—…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Index           =   10
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Œ — ‰Ê⁄ «·›« Ê—…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   12
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   270
         Width           =   1305
      End
   End
   Begin VB.ComboBox DCboCashType 
      Height          =   315
      Left            =   3780
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1785
      Width           =   2655
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   8145
      _cx             =   14367
      _cy             =   1032
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
      BackColor       =   12648447
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "     „ÊÌ· Ê«” ⁄«÷… «·Œ“‰ Ê «·⁄Âœ  "
      Align           =   0
      AutoSizeChildren=   0
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
      FrameColor      =   8454143
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   2940
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   -120
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   3420
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   -120
         Visible         =   0   'False
         Width           =   495
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1155
         TabIndex        =   2
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmPayments1.frx":19F1
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
         Left            =   90
         TabIndex        =   3
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmPayments1.frx":1D8B
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
         Left            =   1680
         TabIndex        =   4
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmPayments1.frx":2125
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
         Left            =   615
         TabIndex        =   5
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmPayments1.frx":24BF
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin MSAdodcLib.Adodc numbering 
         Height          =   585
         Left            =   4560
         Top             =   0
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1032
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   " Õ—Ìﬂ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc detect_no 
         Height          =   585
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1032
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   " Õ—Ìﬂ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   2280
         Picture         =   "FrmPayments1.frx":2859
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   4830
      TabIndex        =   9
      Top             =   6900
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   7110
      TabIndex        =   24
      Top             =   7290
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      Left            =   6225
      TabIndex        =   25
      Top             =   7290
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      Left            =   5340
      TabIndex        =   26
      Top             =   7290
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      Left            =   4455
      TabIndex        =   27
      Top             =   7290
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      Left            =   3570
      TabIndex        =   28
      Top             =   7290
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      Left            =   30
      TabIndex        =   29
      Top             =   7290
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      Height          =   375
      Left            =   915
      TabIndex        =   30
      Top             =   7290
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      Height          =   375
      Index           =   5
      Left            =   2685
      TabIndex        =   31
      Top             =   7290
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      Height          =   375
      Index           =   7
      Left            =   1830
      TabIndex        =   32
      Top             =   7290
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   8
      Left            =   9000
      TabIndex        =   33
      Top             =   6960
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "«·⁄—÷ «·ÃœÊ·Ï"
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
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   9120
      TabIndex        =   56
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«ŸÂ«— «·«ﬁ”«ÿ"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPayments1.frx":64C1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Left            =   9120
      TabIndex        =   57
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "«ŸÂ«— ”‰œ «·„œÌÊ‰Ì…"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPayments1.frx":64DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DCPROJECT 
      Height          =   315
      Left            =   8400
      TabIndex        =   58
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcCostCenter 
      Bindings        =   "FrmPayments1.frx":64F9
      Height          =   315
      Left            =   9480
      TabIndex        =   68
      Top             =   4920
      Width           =   2535
      _ExtentX        =   4471
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   10
      Left            =   4800
      TabIndex        =   71
      Top             =   7680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·‘Ìﬂ"
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
      Index           =   9
      Left            =   3600
      TabIndex        =   72
      Top             =   7680
      Width           =   975
      _ExtentX        =   1720
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
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   2400
      TabIndex        =   81
      Top             =   7680
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   661
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
   Begin MSComCtl2.DTPicker XPDtbTrans 
      Height          =   315
      Left            =   5010
      TabIndex        =   83
      Top             =   960
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      Format          =   159383553
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DcBranch 
      Height          =   315
      Left            =   3780
      TabIndex        =   84
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Accredit 
      Height          =   345
      Left            =   0
      TabIndex        =   125
      Top             =   7680
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   609
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
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
      Height          =   435
      Index           =   39
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   80
      Top             =   8160
      Width           =   7155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄"
      Height          =   195
      Index           =   40
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   78
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ì „  „ÊÌ· Ê «” ⁄«÷… «·Œ“‰ Ê «·⁄Âœ  ”Ê«¡ «· „ÊÌ· ‰ﬁœ« «Ê »‘Ìﬂ«  "
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
      Height          =   540
      Index           =   38
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "„·«ÕŸ… Â«„…:-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   37
      Left            =   2310
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
      Height          =   255
      Left            =   12000
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "œ›⁄Â „ﬁœ„Â"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   35
      Left            =   8730
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   3015
      Width           =   1395
   End
   Begin VB.Label lblsqlstring 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   135
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„‘—Ê⁄"
      Height          =   285
      Index           =   33
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Left            =   1830
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   6930
      Width           =   825
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   6930
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   255
      Index           =   6
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   6930
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   255
      Index           =   7
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   6930
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   300
      Index           =   8
      Left            =   6780
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   6930
      Width           =   1140
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·„œ›Ê⁄« "
      Height          =   285
      Index           =   0
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1800
      Width           =   1485
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   285
      Index           =   1
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   975
      Width           =   1485
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·⁄„·Ì…"
      Height          =   285
      Index           =   4
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   630
      Width           =   1485
   End
End
Attribute VB_Name = "FrmPayments1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim Line1 As Double
Dim Line2 As Double
Dim departement_name As Integer
Dim numbering_type As Integer
Dim Balance As String

Dim balanceString As String

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
If val(XPTxtID.Text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
 
    SendTopost Me.Name, "Notes", "NoteID", 0, val(Dcbranch.BoundText), val(XPTxtID.Text), TxtNoteSerial1.Text
  rs.Resync
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
    fillapprovData
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
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
Accredit.Enabled = False
Else
Accredit.Enabled = True
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
End If
 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.Rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label11.Caption = "Approved"
                                 End If
                            Label11.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label11.Caption = "Currently required Approve"
                            End If
                 Label11.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 GRID2.Rows = 1
    End If
RsDetails.Close

End Function
Private Sub ALLButton1_Click()

    If IsNumeric(Me.DBCboClientName.BoundText) Then
      End If

End Sub

Private Sub ALLButton2_Click()

    If IsNumeric(Me.DBCboClientName.BoundText) Then
   
    End If

End Sub

Private Sub ALLButton3_Click()
    lblsqlstring.Caption = ""
    FrmPaymentTime2.show
    FrmPaymentTime2.lblcusid = DBCboClientName.BoundText
    FrmPaymentTime2.LblValue = val(XPTxtVal.Text)
End Sub

Private Sub CboPayMentType_Change()

    If Me.TxtModFlg.Text = "E" Then
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        Me.DcboBox.Text = ""
    End If

    If Me.CboPaymentType.ListIndex = 0 Then
        Me.lbl(9).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
    ElseIf Me.CboPaymentType.ListIndex = 1 Or Me.CboPaymentType.ListIndex = 3 Then
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
    ElseIf Me.CboPaymentType.ListIndex = 2 Then
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = True
        Me.lbl(16).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True

        'DcboBankName.BoundText = ""
        'TxtChequeNumber.text = ""
        'Frame3.Enabled = True
        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(16).Caption = "—ﬁ„ «·ÕÊ«·…  "
            lbl(17).Caption = " «—ÌŒÂ«"
        Else
            lbl(16).Caption = "Trans No "
            lbl(17).Caption = "Date"
 
        End If

    Else
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(15).Enabled = False
        Me.lbl(16).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub ChkTrans_Click()
    Me.lbl(10).Enabled = ChkTrans.value
    Me.lbl(12).Enabled = ChkTrans.value
    Me.CboTrans.Enabled = ChkTrans.value
    Me.TxtTransID.Enabled = ChkTrans.value
    Me.TxtTransSerial.Enabled = ChkTrans.value
    Me.CmdSearchTrans.Enabled = ChkTrans.value
    Me.CmdOpenTrans.Enabled = ChkTrans.value
End Sub

Function sand_numbering() As String
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    Dim auto_sanad_no As String
    Dim NO As Integer
    auto_sanad_no = ""
    departement_name = 1
 
    connection_string = Cn.ConnectionString
    numbering.ConnectionString = connection_string
    numbering.CommandType = adCmdText
    numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=4"
    numbering.Refresh

    If numbering.Recordset.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = numbering.Recordset.Fields!numbering_id
        start_at = numbering.Recordset.Fields!start_at
        end_at = numbering.Recordset.Fields!end_at

    End If

    If numbering_type = 1 Then
        detect_no.ConnectionString = connection_string
        detect_no.CommandType = adCmdText
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=5 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
        detect_no.Refresh

        If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
 
            If end_at = 0 Then end_at = detect_no.Recordset.Fields!last_sand_no + 1
 
            If detect_no.Recordset.Fields!last_sand_no >= end_at Then
                sand_numbering = "error"
                Exit Function
            End If
        End If

    Else

        If numbering_type = 2 Then
 
            detect_no.ConnectionString = connection_string
            detect_no.CommandType = adCmdText
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=5 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 4, 2)
            'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

            If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)

                If end_at = 0 Then end_at = NO + 1
                If NO >= end_at Then
                    sand_numbering = "error"
                    Exit Function
                End If
            End If

        Else

            If numbering_type = 3 Then
 
                detect_no.ConnectionString = connection_string
                detect_no.CommandType = adCmdText
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  Notes where NoteType=5 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4)
                'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                detect_no.Refresh

                If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)

                    If end_at = 0 Then end_at = NO + 1
                    If NO >= end_at Then
                        sand_numbering = "error"
                        Exit Function
                    End If
                End If
 
            End If
 
        End If
    End If

    If detect_no.Recordset.RecordCount = 0 Or IsNull(detect_no.Recordset.Fields!last_sand_no) Then

        If numbering_type = 0 Then
            ' auto_sanad_no = 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = start_at
            Else
                
                If numbering_type = 2 Then
                    auto_sanad_no = mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = mId(Format$(XPDtbTrans.value, "dd/mm/yyyy"), 7, 4) & start_at

                    End If
                End If
            End If
        End If

    Else

        If numbering_type = 0 Then
            'auto_sanad_no = x + 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = detect_no.Recordset.Fields!last_sand_no + 1
            Else
                
                If numbering_type = 2 Then
                    '  If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) Then
                    ' no = 1
                    '  auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & "1"
                    '  Else
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (NO + 1)
                    '  End If
                      
                Else

                    If numbering_type = 3 Then
                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                        'no = 1
                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                        '    Else
                        NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                        auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (NO + 1)

                        '    End If

                    End If
                End If
            End If
        End If

    End If

    sand_numbering = auto_sanad_no

    'MsgBox auto_sanad_no

End Function

Private Sub Cmd_Click(Index As Integer)
    Dim cNoteReport As ClsNotesReports
    Dim Msg As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If SystemOptions.SysRegisterState = DemoRun Then
                If Not rs Is Nothing Then
                    If Not (rs.BOF Or rs.EOF) Then
                        If rs.RecordCount >= 25 Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "›Ï «·‰”Œ… «· Ã—Ì»Ì… ·«Ì„ﬂ‰  ”ÃÌ· «ﬂÀ— „‰ 25 ⁄„·Ì… ﬁ»÷ «Ê œ›⁄"
                            Else
                            Msg = "In the demo version can not be more than 25 motion recording"
                            End If
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            Exit Sub
                        End If
                    End If
                End If
            End If

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.Text = "N"
            XPTxtID.Text = CStr(new_id("Notes", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=5"))
            Me.DCboUserName.BoundText = user_id
            XPDtbTrans.SetFocus
            Text1.Text = setfoxy
            Me.Dcbranch.BoundText = Current_branch
             Accredit.Caption = ""
             GRID2.Clear flexClearScrollable, flexClearEverything
              GRID2.Rows = 1
        Case 1
    
If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If SystemOptions.banks_Accounts3 = True Then
                If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–… «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
                    Else
                      Msg = " You can not modify this process"
                    Msg = Msg & CHR(13) & "There repayment checks process "
                    
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
           ' Me.DCboUserName.BoundText = user_id
            CuurentLogdata
        
        Case 2
If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If Trim(Dcbranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Dcbranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.Dcbranch.BoundText

            '             If Me.TxtModFlg.text = "N" Then
             
            '             End If
 
            SaveData

        Case 3
            Undo

        Case 4

If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If SystemOptions.banks_Accounts3 = True Then
                If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–… «·⁄„·Ì…"
                    Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
                    Else
                        Msg = " You can not Delete this process"
                    Msg = Msg & CHR(13) & "There repayment checks process "
                    
                    End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If
    
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 50
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                print_report Me.TxtNoteSerial.Text
        
                '     Set cNoteReport = New ClsNotesReports
                '     cNoteReport.PrintReceipt Val(Me.XPTxtID.text), WindowTarget
                '     Set cNoteReport = Nothing
            End If

        Case 9
   
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc Me.TxtNoteSerial.Text, , 200
   
        Case 10
   
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
        
            print_Cheque TxtChequeNumber.Text, get_Cheque_report_no(val(DcboBankName.BoundText)), TxtNoteSerial.Text
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Function print_report(Optional NoteSerial As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From EXPENSES_ORDER2  where NoteSerial='" & NoteSerial & "'"
 
    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
    'End If
    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
    'End If

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "Expenses_order3.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "Expenses_order3.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
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
        xReport.ParameterFields(5).AddCurrentValue DcboDebitSide.Text
   
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
        xReport.ParameterFields(5).AddCurrentValue DcboDebitSide.Text
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
                   ''//////
   Dim xLogo As CRAXDRT.OLEObject
   Dim SqlT As String
   Dim I As Integer
   Dim EmpIDD As Long
   Dim xWidth As Integer
   Dim Rs4 As ADODB.Recordset
   Set Rs4 = New ADODB.Recordset
  SqlT = " SELECT        TOP (100) PERCENT dbo.TblUsers.Empid"
  SqlT = SqlT + "    FROM            dbo.ApprovalData INNER JOIN"
  SqlT = SqlT + "                      dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
  SqlT = SqlT + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.Text) & ") AND (NOT (ApprovDate IS NULL)) AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
  SqlT = SqlT & " ORDER BY levelorder"
  Rs4.Open SqlT, Cn, adOpenStatic, adLockOptimistic, adCmdText
  xWidth = 300
  For I = 1 To Rs4.RecordCount
  EmpIDD = IIf(IsNull(Rs4("Empid").value), 0, Rs4("Empid").value)
            If Dir(App.path & "\" & SystemOptions.ImagesPath & "\sign" & EmpIDD & ".JPG") <> "" Then
          
    
        
            Set xLogo = xReport.Areas(1).Sections(1).AddPictureObject(App.path & "\" & SystemOptions.ImagesPath & "\sign" & EmpIDD & ".JPG", xWidth, 1700)
            xLogo.Width = 800
            xLogo.Height = 400
            xLogo.backcolor = vbWhite
            xLogo.BorderColor = 255
            xLogo.CloseAtPageBreak = True
           xWidth = xWidth + 1000
          End If
        Rs4.MoveNext
    Next I
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub CmdAttach_Click()
     On Error Resume Next
           If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtNoteSerial1, "0712201402"

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdSearchTrans_Click()
    Dim Msg As String

    If Me.CboTrans.ListIndex = -1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·Õ—ﬂ… «·„—«œ «·»ÕÀ ⁄‰Â«..."
        Else
        Msg = "Please Select Type"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboTrans.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If Me.CboTrans.ListIndex = 0 Then
        ' ›« Ê—… „‘ —Ì« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = PurchaseTransaction
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPaymentType.ListIndex = 1
        FrmBuySearch.CboPaymentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… ‘—«¡"
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show
    ElseIf Me.CboTrans.ListIndex = 1 Then
        '›« Ê—… „— Ã⁄ „»Ì⁄« 
        Load FrmBuySearch
        FrmBuySearch.DealingForm = ReturnSalling
        Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
        FrmBuySearch.CboPaymentType.ListIndex = 1
        FrmBuySearch.CboPaymentType.Enabled = False
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ „— Ã⁄ «·„»Ì⁄« "
        FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
        FrmBuySearch.show vbModal
    End If

End Sub

Private Sub DBCboClientName_Change()
 
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1 Then
            Me.DcboDebitSide.BoundText = DBCboClientName.BoundText
        Else
            Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
        End If
    End If

End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
               FrmExpensesSearch.RetrunType = 2508
                FrmExpensesSearch.Indx = 2
                FrmExpensesSearch.Caption = Me.Caption
                FrmExpensesSearch.show
End If

End Sub

Private Sub DcboBankName_Click(Area As Integer)
    On Error Resume Next

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '    Me.DcboCreditSide.BoundText = "a2a3a2"
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If SystemOptions.banks_Accounts3 = True Then
            Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code2")
        Else
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
        
        If CboPaymentType.ListIndex = 2 Or CboPaymentType.ListIndex = 3 Then
                     
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If
    End If

End Sub

Function saveChequeBoxContents1(NoteID As Double)

    If SystemOptions.banks_Accounts3 = False Then Exit Function
    Dim I As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords
 
 '   rs.Open "TblChecqueBoxContent1", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  StrSQL = "SELECT     * from dbo.TblChecqueBoxContent1 Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    If CboPaymentType.ListIndex = 1 Then
        rs.AddNew
        rs("noteid").value = NoteID
     
        rs("RecordDate").value = XPDtbTrans.value
        rs("DueDate").value = DtpChequeDueDate.value
        rs("BankID").value = val(DcboBankName.BoundText)
        rs("BankName").value = DcboBankName.Text
        
        rs("ChequeNo").value = TxtChequeNumber.Text
        rs("ChequeValue").value = val(XPTxtVal.Text)
    
        rs("Remarks").value = Me.DcboDebitSide.Text
        rs("Payed").value = 0
       
        rs("DepitAccount").value = (DcboDebitSide.BoundText)
        rs.update
    End If

    rs.Close
End Function

Private Sub DcboBox_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub

Private Sub DCboCashType_Change()

    Dim StrSQL As String
    Dim intDef As Integer
    Dim Dcombos As ClsDataCombos
    Dim My_SQL As String

    If SystemOptions.UserInterface = EnglishInterface Then
        lbl(3).Caption = "Name"
    Else
        lbl(3).Caption = "«·«”„"
    End If
        
    On Error GoTo ErrTrap

    Select Case DCboCashType.ListIndex + 5

        Case 0
            Set Dcombos = New ClsDataCombos
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
            ChkTrans.Visible = True
            Fra(0).Visible = True

        Case 1
            Set Dcombos = New ClsDataCombos
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
            ChkTrans.Visible = True
            Fra(0).Visible = True

        Case 2
            Set Dcombos = New ClsDataCombos
            Dcombos.GetPersons Me.DBCboClientName
            ChkTrans.Visible = False
            Fra(0).Visible = False

        Case 3
            Fra(0).Visible = True

            If SystemOptions.UserInterface = EnglishInterface Then
                lbl(3).Caption = "Project"
            Else
                lbl(3).Caption = "«·„‘—Ê⁄"
            End If
        
            My_SQL = "  select expanses_account,Project_name from projects where not(expanses_account is null)" '  where  Account_code like'" & Account_Code_dynamic & "%' and last_account=1"
            fill_combo Me.DBCboClientName, My_SQL

        Case 4
            Frame2.Enabled = True
            My_SQL = "  select Account_Code,Account_Name from ACCOUNTS where last_account=1"
            fill_combo Me.DBCboClientName, My_SQL
            Option4.value = True
      
        Case 5
            Set Dcombos = New ClsDataCombos
            Dcombos.GetBoxAccounts Me.DBCboClientName, 1
     
            ' My_SQL = "  select Account_Code,BoxName from TblBoxesData where Type=1"
            'fill_combo Me.DBCboClientName, My_SQL
      
        Case 6
       
            '      My_SQL = "  select Account_Code,BoxName from TblBoxesData where Type=0"
            '      fill_combo Me.DBCboClientName, My_SQL
         
            Set Dcombos = New ClsDataCombos
            Dcombos.GetBoxAccounts Me.DBCboClientName, 0
        
    End Select

    cSearchDcbo.Refresh
    Set Dcombos = Nothing
    Exit Sub
ErrTrap:

End Sub

Private Sub DCboCashType_Click()
    DCboCashType_Change
End Sub

Private Sub ChangeLang()
CmdAttach.Caption = "Attachments"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Me.Shape1.Visible = False
    lbl(37).Visible = False
    lbl(38).Visible = False
    Frame1.Caption = "Options"
    Option3.Caption = "Adv. Payment"
    Option2.Caption = "Select Invoice"
    ALLButton3.Caption = "Select"
    lbl(22).Caption = "Current Week"
    lbl(35).Caption = "Adv. Payment"
    Label8.Caption = "General C.C."
    lbl(36).Caption = "General Des"
    Cmd(9).Caption = "GL Print"
    Cmd(10).Caption = "Cheque Print"
    Frame2.Caption = "Employee"
    Option4.Caption = "Salary"
    Option5.Caption = "Advance"

    lbl(18).Visible = False
    ALLButton1.Caption = "Installment view"
    ALLButton2.Caption = "debt Voucher"
    Me.Caption = " Boxes  Recharge"
    
    C1Elastic1.Caption = Me.Caption
    lbl(4).Caption = "Opr Code"
    lbl(1).Caption = "Date"
    lbl(40).Caption = "Branch"

    lbl(0).Caption = "Type"
    lbl(3).Caption = "Name"
    lbl(2).Caption = "Value"
    lbl(14).Caption = "Payemnt Method"
    lbl(9).Caption = "Box Name"
    lbl(15).Caption = "Bank Name"
    lbl(16).Caption = "Cheque #"
    lbl(17).Caption = "Cheque date"
    lbl(34).Caption = "Due To"
    lbl(5).Caption = "Note"
    ChkTrans.Caption = "From bill"
    lbl(12).Caption = "Bill type"
    lbl(10).Caption = "Bill #"
    lbl(13).Caption = "Current Balance"
    FraInfo.Caption = "Information"
    lbl(22).Caption = "Current Week"

    lbl(23).Caption = "Today Payments "
    lbl(27).Caption = "Cash"
    lbl(28).Caption = "Cheque"

    lbl(19).Caption = "Week Payments "

    lbl(21).Caption = "Cash"
    lbl(24).Caption = "Cheque"

    lbl(20).Caption = "Month Payments "

    lbl(25).Caption = "Cash"
    lbl(26).Caption = "Cheque"
    Fra(1).Caption = "GL"

    lbl(30).Caption = "GL#"
    lbl(29).Caption = "Interval"

    lbl(31).Caption = "Depit"
    lbl(32).Caption = "Credit"
    Cmd(8).Caption = "Table view"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Current "
    lbl(6).Caption = "Records Count "

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    DCboCashType.Clear
    'DCboCashType.AddItem "To Customer"
    'DCboCashType.AddItem "To Vendor"
    'DCboCashType.AddItem "sub-contractor"
    'DCboCashType.AddItem "To Project"
    'DCboCashType.AddItem "To Employee"
    DCboCashType.AddItem "Bety Cash"
    DCboCashType.AddItem "Box Recharge"

    Option4.Caption = "Salary"
    Option5.Caption = "Advance"

    With Me.CboPaymentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
        .AddItem "Transfer"
        .AddItem "P Cheque"
    
    End With

    With Me.CboTrans
        .Clear
        .AddItem "Purchase invoice"
        .AddItem "Returned sales"
    End With

End Sub

Private Sub DcboDebitSide_Change()
    WriteCustomerBalPublic Me.DcboDebitSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub Form_Load()
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500

    ScreenNameArabic = "«” ⁄«÷… ⁄ÂœÂ Ê „ÊÌ· Œ“Ì‰…"
    ScreenNameEnglish = "Bt Cash"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 50

    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos

    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
 
    fill_combo Me.DcCostCenter, StrSQL

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    'Resize_Form Me

    AddTip
    'DCboCashType.AddItem "≈·Ï ⁄„Ì·"
    'DCboCashType.AddItem "≈·Ï „Ê—œ"
    'DCboCashType.AddItem "„ﬁ«Ê· »«ÿ‰"
    'DCboCashType.AddItem "„‘—Ê⁄"
    'DCboCashType.AddItem "„ÊŸ›"
    DCboCashType.AddItem "«” ⁄«÷Â ⁄ÂœÂ"
    DCboCashType.AddItem " „ÊÌ· Œ“Ì‰…"

    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName

    Dcombos.GetBranches Me.Dcbranch

    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    With Me.CboPaymentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "ÕÊ«·…"
        .AddItem "‘Ìﬂ „”œœ"
    
    End With

    With Me.CboTrans
        .Clear
        .AddItem "›« Ê—… „‘ —Ì« "
        .AddItem "„— Ã⁄ „»Ì⁄« "
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DBCboClientName
    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide

    Set rs = New ADODB.Recordset
    'StrSQL = "select * From Notes where  NoteType=50  or (NoteType=5 and (CashingType=5 OR CashingType=6))order by NoteID"
    StrSQL = "select * From Notes where  NoteType=50   "

    If SystemOptions.usertype <> UserAdminAll Then
      '  StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
    StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
    
StrSQL = StrSQL & "order by NoteID "

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    SetDtpickerDate XPDtbTrans
    SetDtpickerDate Me.DtpChequeDueDate
    ChkTrans.value = Unchecked
    ChkTrans_Click

    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"
    WriteInfo
    Dim My_SQL As String

    'My_SQL = "  select account_no,account_name from projects  where not (account_no is null)"
    My_SQL = "  select expanses_account,Project_name from projects where not(expanses_account is null)" '  where  Account_code like'" & Account_Code_dynamic & "%' and last_account=1"
    fill_combo dcproject, My_SQL

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, 50

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
    Exit Sub
ErrTrap:
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption



End Sub

Private Sub LblLink_Click()
    Dim FirstPeriod As Date
    getFirstPeriodDateInthisYear FirstPeriod
    ShowReport DcboDebitSide.BoundText, DcboDebitSide.Text, FirstPeriod, Date

End Sub

Private Sub LblLink_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              Y As Single)

    If SystemOptions.UserInterface = ArabicInterface Then
        LblLink.ToolTipText = "—’Ìœ «·ÿ—› «·„œÌ‰:" & WriteNo(Balance, 0, True)
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        LblLink.ToolTipText = "Depit Balance:" & WriteNo(Balance, 0, True)
    End If

End Sub

Private Sub Option1_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

End Sub

Private Sub Option2_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

End Sub

Private Sub Option3_Click()

    If Option2.value = True Then
        ALLButton3.Enabled = True
    Else
        ALLButton3.Enabled = False
    End If

End Sub

Private Sub Option4_Click()
    '  Dim My_SQL As String
    '  My_SQL = "  select Account_Code1,Emp_Name from TblEmployee"
    '  fill_combo Me.DBCboClientName, My_SQL
    '  If txt_general_des.text = "" And Me.TxtModFlg <> "R" Then
    '      txt_general_des.text = Option4.Caption
    '  End If
End Sub

Private Sub Option5_Click()
    '  Dim My_SQL As String
    '  My_SQL = "  select Account_Code,Emp_Name from TblEmployee"
    '  fill_combo Me.DBCboClientName, My_SQL
    '  If txt_general_des.text = "" And Me.TxtModFlg <> "R" Then
    '      txt_general_des.text = Option5.Caption
    '  End If
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            Frame2.Enabled = False

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Box Recharge "
            Else
                Me.Caption = " „ÊÌ· «·Œ“Ì‰… Ê«” ⁄«÷… «·⁄Âœ"
     
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            DBCboClientName.Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            XPTxtVal.locked = True
            XPDtbTrans.Enabled = False
            XPMTxtRemarks.locked = True
            DBCboClientName.locked = True
            DcboBox.locked = True
            DCboCashType.locked = True
            Me.CboPaymentType.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

            Fra(0).Enabled = False
            ChkTrans.Enabled = False

        Case "N"

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Box Recharge (New)"
            Else
                Me.Caption = "  „ÊÌ· «·Œ“Ì‰… Ê«” ⁄«÷… «·⁄Âœ"
        
            End If

            DBCboClientName.Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
            XPDtbTrans.Enabled = True
            XPTxtVal.locked = False
            XPMTxtRemarks.locked = False
            DBCboClientName.locked = False
            DcboBox.locked = False
            Me.CboPaymentType.locked = False
            XPDtbTrans.value = Date
            DCboCashType.locked = False
            DBCboClientName.locked = False
            DCboCashType.ListIndex = 1
            Fra(0).Enabled = True
            ChkTrans.Enabled = True

        Case "E"

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Box Recharge (Edit)"
            Else
                Me.Caption = "  „ÊÌ· «·Œ“Ì‰… Ê«” ⁄«÷… «·⁄Âœ(  ⁄œÌ· )"
        
            End If

            DBCboClientName.Enabled = True
    
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
            XPTxtVal.locked = False
            XPDtbTrans.Enabled = True
            DcboBox.locked = False
            XPMTxtRemarks.locked = False
            DBCboClientName.locked = False
            DCboCashType.locked = False
            DBCboClientName.locked = False
            Me.CboPaymentType.locked = False
            Fra(0).Enabled = True
            ChkTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtTransID_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If Me.TxtTransID.Text <> "" Then
            If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                Me.TxtTransSerial.Text = GetTransIDSerial(1, val(Me.TxtTransID.Text))
            Else
                Me.TxtTransSerial.Text = ""
            End If
        End If
    End If

End Sub

Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTransSerial.Text, 1)
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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim RsDev As ADODB.Recordset
    Dim I As Integer

    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    End If

    Me.Text1.Text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    XPTxtID.Text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    Me.oldtxtNoteSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(39).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    txt_general_des.Text = IIf(IsNull(rs("general_des_notes").value), "", rs("general_des_notes").value)

    txtperson.Text = IIf(IsNull(rs("person").value), "", rs("person").value)

    XPTxtVal.Text = IIf(IsNull(rs("Note_Value").value), "", Trim(rs("Note_Value").value))
    dcproject.BoundText = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))
    XPMTxtRemarks.Text = IIf(IsNull(rs("Remark").value), "", Trim(rs("Remark").value))

    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    DCboCashType.ListIndex = IIf(IsNull(rs("CashingType").value), -1, rs("CashingType").value - 5)

    Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)

    If DCboCashType.ListIndex = 3 Then
        DBCboClientName.BoundText = IIf(IsNull(rs("projectAccountCode").value), 0, rs("projectAccountCode").value)

    ElseIf DCboCashType.ListIndex = 4 Then
        DBCboClientName.BoundText = IIf(IsNull(rs("EmpAccountCode").value), 0, rs("EmpAccountCode").value)

    ElseIf DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1 Then
        DBCboClientName.BoundText = IIf(IsNull(rs("BTCashAccountcode").value), 0, rs("BTCashAccountcode").value)
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
    End If

    '---------------------------------------------------------------------------
    If IsNull(rs("salary_or_advance").value) Then
        Option4.value = False: Option5.value = False
    ElseIf (rs("salary_or_advance").value) = 0 Then
        Option4.value = True
    ElseIf (rs("salary_or_advance").value) = 1 Then
        Option4.value = False
    End If

    If IsNull(rs("NoteCashingType").value) Then
        Me.CboPaymentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPaymentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
    ElseIf rs("NoteCashingType").value = 1 Then
        Me.CboPaymentType.ListIndex = 1
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
 
    ElseIf rs("NoteCashingType").value = 2 Then
        Me.CboPaymentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value

    ElseIf rs("NoteCashingType").value = 3 Then
        Me.CboPaymentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value

    End If

    CboPayMentType_Change

    '-----------------------------------------------------------------------------
    If Not IsNull(rs("Transaction_ID").value) Then
        Me.ChkTrans.value = vbChecked
        'Me.ChkTrans.Enabled = True
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select * From Transactions Where Transaction_ID=" & rs("Transaction_ID").value
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            Me.TxtTransID.Text = RsTemp("Transaction_ID").value
            Me.TxtTransSerial.Text = IIf(IsNull(RsTemp("Transaction_Serial").value), "", RsTemp("Transaction_Serial").value)

            If Not (IsNull(RsTemp("Transaction_Type").value)) Then
                If RsTemp("Transaction_Type").value = 9 Then
                    Me.CboTrans.ListIndex = 1
                ElseIf RsTemp("Transaction_Type").value = 1 Then
                    Me.CboTrans.ListIndex = 0
                End If
            End If
        End If

    Else
        Me.ChkTrans.value = vbUnchecked
        Me.CboTrans.ListIndex = -1
        Me.TxtTransID.Text = ""
        Me.TxtTransSerial.Text = ""
    End If

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.XPTxtID.Text)
        StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lbl(11).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For I = 1 To RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                End If

                RsDev.MoveNext
            Next I

        End If
    End If
fillapprovData
    '-----------------------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim StrTemp As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
Dim maxvalue As Double
Dim x As Integer
    On Error GoTo ErrTrap
     Dim Posted As Integer
            If CheckAprroveScreen(Me.Name) = True Then
            Posted = 1
            Else
            Posted = 0
            End If
    If Me.TxtModFlg.Text <> "R" Then
        
If CheckBoxmaxVaue((DBCboClientName.BoundText), val(XPTxtVal.Text), maxvalue) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
x = MsgBox("«–«  „  Â–… «·⁄„·Ì… ﬁœ   ⁄œÌ «ﬁ’Ì ﬁÌ„… ··⁄Âœ… ÊÂÌ " & maxvalue & CHR(13) & "  «· ﬂ„·… ⁄·Ì «Ì Õ«·", vbInformation + vbYesNo)
Else
x = MsgBox("If this process will exceed the maximum value for the custody of the " & maxvalue & CHR(13) & " Supplement anyway", vbInformation + vbYesNo)

End If
If x = vbNo Then

    Exit Sub
End If


End If
        
        If DCboCashType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·„œ›Ê⁄«  "
            Else
            Msg = "Please Select Type Payment"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCboCashType.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If DBCboClientName.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» «Œ Ì«— «”„ «·⁄„Ì· √Ê «·„Ê—œ"
            Else
            Msg = "Please Select Name"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If XPTxtVal.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·„œ›Ê⁄«  "
            Else
            Msg = "Please Enter Value"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtVal.SetFocus
            Exit Sub
        End If

        If Not IsNumeric(XPTxtVal.Text) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ﬁÌ„… «·„œ›Ê⁄«  ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            Else
            Msg = "The value must be digital"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtVal.SetFocus
            SelectText XPTxtVal
            Exit Sub
        End If

        If Me.CboPaymentType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
            Else
            Msg = "Please Select Type of Payment"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboPaymentType.SetFocus
            Exit Sub
        End If

        If Me.CboPaymentType.ListIndex = 0 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
                Else
                Msg = "Please Select  safe"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBox.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPaymentType.ListIndex = 1 Or Me.CboPaymentType.ListIndex = 3 Then

            If Me.DcboBankName.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                Msg = "Please Select Bank"
                End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBankName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.Text) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                Else
                Msg = "You must write the check number"
                End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If

            '      If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
            '          Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
            '          MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '          DtpChequeDueDate.SetFocus
            '          SendKeys "{F4}"
            '          Exit Sub
            '      End If
        ElseIf Me.CboPaymentType.ListIndex = 2 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
                                        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBankName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.Text) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·ÕÊ«·…...!!"
                Else
                    Msg = "Enter Cheque No:...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            
            End If

        End If
    
        If Me.TxtModFlg.Text = "N" Then
            If Me.CboPaymentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.Text), XPDtbTrans.value) = False Then
                        Exit Sub
                    End If
                End If
            End If

        ElseIf Me.TxtModFlg.Text = "E" Then

            If Me.CboPaymentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.Text), XPDtbTrans.value, , , val(Me.XPTxtID.Text)) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        If Me.ChkTrans.value = vbChecked Then
            If Me.CboTrans.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·›« Ê—…..!!!"
                Else
                Msg = "Please select the invoice type"
                End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                CboTrans.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

            If Trim(Me.TxtTransSerial.Text) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡ ≈œŒ«· —ﬁ„ «·›« Ê—…..!!!"
                Else
                Msg = "Please select the invoice number"
                End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtTransSerial.SetFocus
                Exit Sub
            Else

                If Me.CboTrans.ListIndex = 0 Then
                    If Me.TxtTransID.Text = "" Then
                        StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.Text), 1, Me.DBCboClientName.BoundText)
                    Else
                        StrTemp = Me.TxtTransID.Text
                    End If

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If

                ElseIf Me.CboTrans.ListIndex = 1 Then
                    StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.Text), 9)

                    If CheckDebitTrans(val(StrTemp)) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If
my_branch = val(Me.Dcbranch.BoundText)

        If TxtNoteSerial.Text = "" Then
            If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                Else
                MsgBox "Can't Add new G.E Codeing Exceed": Exit Sub
                End If
            Else
                       
                If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                    MsgBox "Define G.E coding": Exit Sub
                    End If
                Else
                    TxtNoteSerial.Text = Notes_coding(val(my_branch), XPDtbTrans.value)
                End If
            End If
        End If
        
        If TxtNoteSerial1.Text = "" Then
            If Voucher_coding(val(my_branch), XPDtbTrans.value, 5, 50) = "error" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  „ÊÌ· Ê «” ⁄«÷… ÃœÌœ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                MsgBox "Can't Add new doc codeing Exceed": Exit Sub
                End If
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbTrans.value, 5, 50) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                    MsgBox "Defin voucher coding": Exit Sub
                    End If
                Else
                    TxtNoteSerial1.Text = Voucher_coding(val(my_branch), XPDtbTrans.value, 5, 50)
                End If
            End If
        End If
    
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.Text = "N" Then
            XPTxtID.Text = CStr(new_id("Notes", "NoteID", "", True))
            '  Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=5"))
            rs.AddNew
            rs("NoteID").value = val(XPTxtID.Text)
            Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
        ElseIf TxtModFlg.Text = "E" Then
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(XPTxtID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If

        rs("branch_no").value = val(Me.Dcbranch.BoundText)
        rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.Text) = "", Null, Trim(Me.TxtNoteSerial.Text))
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
        rs("Note_Value").value = IIf(XPTxtVal.Text = "", Null, val(XPTxtVal.Text))
        rs("note_value_by_characters").value = IIf(lbl(18).Caption = "", Null, lbl(18).Caption)
     
        rs("Remark").value = IIf(XPMTxtRemarks.Text = "", "", Trim(XPMTxtRemarks.Text))
        rs("general_des_notes").value = IIf(txt_general_des.Text = "", "", Trim(txt_general_des.Text))
    
        If txtperson.Text = "" Then
            txtperson.Text = DBCboClientName.Text
        End If

        rs("person").value = IIf(txtperson.Text = "", "", Trim(txtperson.Text))

        rs("NoteType").value = 50
        rs("NoteDate").value = XPDtbTrans.value
        rs("CashingType").value = IIf(DCboCashType.ListIndex = -1, Null, DCboCashType.ListIndex + 5)
    
        If DCboCashType.ListIndex = 3 Then
            rs("projectAccountCode").value = IIf(DBCboClientName.Text = "", Null, DBCboClientName.BoundText)

        ElseIf DCboCashType.ListIndex = 4 Then
            rs("EmpAccountCode").value = IIf(DBCboClientName.Text = "", Null, DBCboClientName.BoundText)
      
        ElseIf DCboCashType.ListIndex = 0 Or DCboCashType.ListIndex = 1 Then
            rs("BTCashAccountcode").value = IIf(DBCboClientName.Text = "", Null, DBCboClientName.BoundText)
      
        Else
            rs("CusID").value = IIf(DBCboClientName.Text = "", Null, DBCboClientName.BoundText)
 
        End If
    
        If Option4.value = True Then
            rs("salary_or_advance").value = 0
        ElseIf Option5.value = True Then
            rs("salary_or_advance").value = 1
        Else
            rs("salary_or_advance").value = Null
        End If
    
        'DcboBox
        If Me.ChkTrans.value = vbChecked Then
            If Me.CboTrans.ListIndex = 0 Or Me.CboTrans.ListIndex = 1 Then
                rs("Transaction_ID").value = val(Me.TxtTransID.Text)
            End If

        Else
            rs("Transaction_ID").value = Null
        End If

        If Me.CboPaymentType.ListIndex = 0 Then
            rs("BoxID").value = val(DcboBox.BoundText)
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("NoteCashingType").value = 0
        ElseIf Me.CboPaymentType.ListIndex = 1 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 1
        
        ElseIf Me.CboPaymentType.ListIndex = 2 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 2
        ElseIf Me.CboPaymentType.ListIndex = 3 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 3
        
        End If

        rs("UserID").value = user_id
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        rs("foxy_no").value = val(Text1.Text)
        rs("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        rs("numbering_type1").value = sand_numbering_type(4) '”‰œ «·œ›⁄
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)
        rs("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
    
        rs.update
    
        If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
            Set RsDev = New ADODB.Recordset
            'RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
                     StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
            Line1 = setfoxy_Line
            Line2 = setfoxy_Line
            '«·ÿ—› «·„œÌ‰
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            If Posted = 1 Then
            RsDev("Posted").value = 1
            Else
            RsDev("Posted").value = Null
            End If
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 1
            RsDev("DEV_ID_Line_No1").value = Line1
            RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.Text)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text
            
            If DCboCashType.ListIndex = 3 Then
                Dim project_id As Integer
                project_id = get_project_id(DBCboClientName.BoundText, "expanses_account")
                RsDev("project_id").value = project_id
                RsDev("Double_Entry_Vouchers_Description").value = "’—› ⁄·Ï „‘—Ê⁄" & DBCboClientName.Text
            End If
            
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            
            RsDev.update
            '«·ÿ—› «·œ«∆‰
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            If Posted = 1 Then
            RsDev("Posted").value = 1
            Else
            RsDev("Posted").value = Null
            End If
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            RsDev("DEV_ID_Line_No").value = 2
            RsDev("Account_Code").value = Me.DcboCreditSide.BoundText
            RsDev("Value").value = val(Me.XPTxtVal.Text)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("DEV_ID_Line_No1").value = Line2
            RsDev("Double_Entry_Vouchers_Description").value = XPMTxtRemarks.Text
            ' RsDev("Double_Entry_Vouchers_Description").value = dcproject.BoundText
            
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
            LblDevID.Caption = LngDevID
            lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
        End If

        saveChequeBoxContents1 (val(XPTxtID.Text))
 
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        CuurentLogdata
   
        Select Case Me.TxtModFlg.Text

            Case "N"
 
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Operation data was saved " & CHR(13)
                    Msg = Msg + "need another operation"
        
                Else
                
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
               
                End If
          
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    MsgBox "Save successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
              
                lbl(39).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
              
        End Select

        TxtModFlg.Text = "R"
fillapprovData
        '«· Ê“Ì⁄ ⁄·Ï „—ﬂ“ «· ﬂ·›… «·⁄«„
        If Me.DcCostCenter.BoundText <> "" Then
            save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.Text, "„œ›Ê⁄« ", Me.XPDtbTrans.value
        End If
        
    End If

    WriteCustomerBalPublic Me.DcboDebitSide.BoundText, Balance, balanceString
    LblLink.Caption = balanceString
    WriteInfo

    If Option1.value = True Then
        FIFO_FUNCTION val(DBCboClientName.BoundText)
    End If
   
    If Option2.value Then
        Distribute_to_bills Me.lblsqlstring, val(DBCboClientName.BoundText)
    End If
   
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
        Msg = "You can not save this data " & CHR(13)
        Msg = Msg + "It has been enter incorrect data " & CHR(13)
        Msg = Msg + "Make sure of the validity of the data and try again"
        
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
     Msg = "Sorry ... error during save " & CHR(13)
    
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim I As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
 
   ' rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  StrSQL = "SELECT   *  from dbo.marakes_taklefa_temp Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    'ÿ—› „œÌ‰
    rs.AddNew
    rs("cost_center_id").value = cost_center_id
    rs("cost_center").value = cost_center
    rs("value").value = XPTxtVal.Text
    rs("depit_or_credit").value = "„œÌ‰"
    rs("opr_id").value = Me.Text1.Text
    rs("kedno").value = Me.Text1.Text
        
    rs("opr_type").value = opr_type
    rs("account_name").value = DcboDebitSide.Text
    rs("account_no").value = DcboDebitSide.BoundText
    rs("line_no").value = Line1
    rs("record_date").value = record_date
    rs.update
    'ÿ—› œ«∆‰
    rs.AddNew
    rs("cost_center_id").value = cost_center_id
    rs("cost_center").value = cost_center
    rs("value").value = XPTxtVal.Text
    rs("depit_or_credit").value = "œ«∆‰"
    rs("opr_id").value = Me.Text1.Text
    rs("kedno").value = Me.Text1.Text
        
    rs("opr_type").value = opr_type
    rs("account_name").value = DcboCreditSide.Text
    rs("account_no").value = DcboCreditSide.BoundText
    rs("line_no").value = Line2
    rs("record_date").value = record_date
    rs.update
 
    rs.Close
End Function

Function FIFO_FUNCTION(CusID As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim I As Integer
    sql = "SELECT CompanyDebitValues.* FROM dbo.CompanyDebitValues() CompanyDebitValues  where   (cusid=" & CusID & " and requiredvalue>0)"

    'Sql = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues  where   (cusid=" & CusID & " and requiredvalue>0)"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Dim total_value As Double
    Dim current_value As Double
    total_value = val(txtAdv_payment_value.Text)
  
    For I = 1 To Rs3.RecordCount

        If total_value > Rs3("requiredvalue") Then
            current_value = Rs3("requiredvalue")
            total_value = total_value - current_value
        
        ElseIf total_value <= Rs3("requiredvalue") Then
            current_value = total_value
            total_value = 0
        ElseIf total_value = 0 Then
            Exit Function
        End If
  
        Add_new_notes Me.XPDtbTrans, 2001, current_value, Rs3("transactionsid").value, CusID, DcboBox.BoundText, 1, val(DCboUserName.BoundText)
  
        Rs3.MoveNext
    Next I

    ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
    txtAdv_payment_value.Text = total_value
    change_adv_payment_value XPTxtID.Text, total_value
    ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
    Rs3.Close

End Function

Function Add_new_notes(NoteDate As Date, NoteType As Integer, Note_Value As Double, Transaction_ID As Integer, CusID As Double, BoxID As Integer, displayed As Integer, UserID As Integer)
    Dim RsDev As New ADODB.Recordset
    Dim StrSQL As String
   ' RsDev.Open "notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    '
        
    RsDev.AddNew
      
    RsDev("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    RsDev("NoteSerial").value = TxtNoteSerial.Text ' CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=2000"))
              
    RsDev("NoteDate").value = NoteDate
    RsDev("NoteType").value = NoteType
           
    RsDev("Note_Value").value = Note_Value
    RsDev("Transaction_ID").value = Transaction_ID
    RsDev("CusID").value = CusID
    RsDev("BoxID").value = BoxID
    RsDev("UserID").value = UserID
    RsDev("displayed").value = 0
           
    RsDev.update

End Function

Function change_adv_payment_value(note_id As Double, value As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim I As Integer

    sql = "SELECT * from notes   where  NoteID=" & note_id
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Rs3("Adv_payment_value").value = value
    Rs3.update
  
End Function

Function Distribute_to_bills(Sql1 As String, CusID As Double)
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim I As Integer

    sql = "SELECT CompanyDebitValues.* FROM dbo.CompanyDebitValues() CompanyDebitValues  where    requiredvalue>0 and " & Sql1

    'Sql = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues  where  requiredvalue>0 and " & Sql1
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function
    Dim total_value As Double
    Dim current_value As Double
    total_value = val(txtAdv_payment_value.Text)
  
    For I = 1 To Rs3.RecordCount

        If total_value > Rs3("requiredvalue") Then
            current_value = Rs3("requiredvalue")
            total_value = total_value - current_value
        
        ElseIf total_value <= Rs3("requiredvalue") Then
            current_value = total_value
            total_value = 0
        ElseIf total_value = 0 Then
            Exit Function
        End If
  
        Add_new_notes Me.XPDtbTrans, 2001, current_value, Rs3("transactionsid").value, CusID, DcboBox.BoundText, 1, val(DCboUserName.BoundText)
        Rs3.MoveNext
    Next I

    txtAdv_payment_value.Text = total_value
    change_adv_payment_value XPTxtID.Text, total_value

    ' If IsNull(Rs3("UserName").value) Then FIFO_FUNCTION = "": Exit Function
  
    ' If Not IsNull(Rs3("UserName").value) Then get_user_name = Rs3("UserName").value: Exit Function
    Rs3.Close
 
End Function

Private Function CheckDebitTrans(LngTransID As Long) As Boolean
Exit Function
    Dim Msg As String
    Dim RsTemp As ADODB.Recordset
    Dim DblCreditNoteValue As Double
    Dim LngDebitNoteID As Long
    Dim StrSQL As String

    CheckDebitTrans = False

    If LngTransID = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê« .. ·« ÊÃœ ›« Ê—… »Â–« «·„”·”· „”Ã·… ›Ï «·»—‰«„Ã..!!!"
        Msg = Msg & CHR(13) & "»—Ã«¡ «· «ﬂœ „‰ «·»Ì«‰«  «·„œŒ·…..!!"
        Else
         Msg = "Sorry There is a bill in this sequence in the program.!!!"
        Msg = Msg & CHR(13) & "Please make sure the data entered..!!"
       
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtTransSerial.SetFocus
        Exit Function
    ElseIf LngTransID <> 0 Then
        Set RsTemp = New ADODB.Recordset
        StrSQL = "Select CusID,PaymentType From Transactions where Transaction_ID=" & LngTransID & ""
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If RsTemp("PaymentType").value = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
                Msg = Msg & CHR(13) & "›« Ê—… ‰ﬁœÌ… ...Ê·«Ì„ﬂ‰  ”ÃÌ· „œ›Ê⁄«  ·Â«"
                Else
                  Msg = "Sorry invoice number " & Trim(Me.TxtTransSerial.Text)
                Msg = Msg & CHR(13) & "Cash bill can not be registered payments "
                
                End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If Me.DBCboClientName.BoundText <> IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ : " & Trim(Me.TxtTransSerial.Text)
                Msg = Msg & CHR(13) & "—ﬁ„ «·›« Ê—… ›Ï «·»—‰«„Ã : " & Me.TxtTransID.Text
                Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· : " & Me.DBCboClientName.Text
                Else
                  Msg = "Sorry invoice number : " & Trim(Me.TxtTransSerial.Text)
                Msg = Msg & CHR(13) & "The invoice number in the program: " & Me.TxtTransID.Text
                Msg = Msg & CHR(13) & "It is not registered with the customer : " & Me.DBCboClientName.Text
                End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtTransSerial.SetFocus
                Exit Function
            End If

            If LngTransID <> val(Me.TxtTransID.Text) Then
                Me.TxtTransID.Text = LngTransID
            End If
        
            DblCreditNoteValue = 0
            StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType, " & "Notes.Note_Value, Notes.NoteID "
            StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & "Notes.Transaction_ID WHERE (Notes.NoteType=1) AND Transactions.Transaction_ID= " & LngTransID & ""
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                LngDebitNoteID = RsTemp("NoteID").value
                DblCreditNoteValue = IIf(IsNull(RsTemp("Note_Value").value), 0, RsTemp("Note_Value").value)
                '«· «ﬂœ „‰ «‰ Â–Â «·›« Ê—… ·Ì”  ·Â« √ﬁ”«ÿ
                'ÕÌÀ «‰ «·√ﬁ”«ÿ ·«Ì„ﬂ‰  Õ’Ì·Â« „‰ Â‰«
                StrSQL = "Select * From InstallMent Where NoteID=" & LngDebitNoteID & ""
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly

                If Not (RsTemp.BOF Or RsTemp.EOF) Then
                    If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "⁄›Ê« .. «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ﬁœ  „  ﬁ”ÌÿÂ«..!!"
                        Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  Õ’Ì· «·√ﬁ”«ÿ „‰ ‘«‘… «·„œ›Ê⁄« "
                        Msg = Msg & CHR(13) & "≈” Œœ„ ‘«‘…  Õ’Ì· «·√ﬁ”«ÿ »œ·« „‰Â«"
                        Else
                         Msg = "Sorry .. futures value of the invoice has been installments..!!"
                        Msg = Msg & CHR(13) & "And you can not collect premiums from the screen of payments"
                        Msg = Msg & CHR(13) & "Use the collection of installments screen"
                       
                        End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Exit Function
                    End If
                End If

            Else
                'LngDebitNoteID
                If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«ÌÊÃœ «Ê—«ﬁ „«·Ì… √Ã·… ⁄·Ï Â–Â «·›« Ê—…..!!"
                Else
                 Msg = "No securities futures on this bill..!!"
                
                End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Function
            End If

            If DblCreditNoteValue < val(Me.XPTxtVal.Text) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ..."
                Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… .. «’€— „‰ «·ﬁÌ„…"
                Msg = Msg & CHR(13) & "«·„—«œ  ”ÃÌ·Â« «·√‰..»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·….!"
                Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                Else
                   Msg = "Sorry ..."
                Msg = Msg & CHR(13) & "Futures value of the invoice is smaller than the value"
                Msg = Msg & CHR(13) & "To be registered, please review the recorded value.!"
                Msg = Msg & CHR(13) & "Note:-"
                Msg = Msg & CHR(13) & "Futures value of the invoice is : " & DblCreditNoteValue
                End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.XPTxtVal.SetFocus
                Exit Function
            End If

            Set RsTemp = New ADODB.Recordset
            StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType," & "Sum(Notes.Note_Value) AS SumNote_Value "
            StrSQL = StrSQL + " FROM Transactions INNER JOIN Notes ON Transactions.Transaction_ID =" & "Notes.Transaction_ID " & " Where ((Notes.NoteType = 5 OR Notes.NoteType = 10) And Transactions.Transaction_ID = " & LngTransID & ")"

            If Me.TxtModFlg.Text = "E" Then
                StrSQL = StrSQL + " And Notes.NoteID <>" & Me.XPTxtID.Text & ""
            End If

            StrSQL = StrSQL + " GROUP BY Transactions.Transaction_ID, Transactions.Transaction_Serial," & "Transactions.Transaction_Type, Transactions.PaymentType "
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If DblCreditNoteValue = RsTemp("SumNote_Value").value Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "⁄›Ê« ...!!!!!" & CHR(13)
                    Msg = Msg & "·ﬁœ  „  ”ÃÌ· „œ›Ê⁄«  √Ê (⁄„· Œ’Ê„«  „ﬂ ”»…) ·Â–Â «·›« Ê—… »„« Ì”«ÊÏ «·ﬁÌ„… «·√Ã·… „‰Â«"
                    Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰   ”ÃÌ· «Ì… „œ›Ê⁄«  ≈÷«›Ì… ⁄·ÌÂ«."
                    Else
                     Msg = "Sorry ...!!!!!" & CHR(13)
                    Msg = Msg & "I've been recording payments or discounts work acquired for this bill, including the value equal to its futures"
                    Msg = Msg & CHR(13) & "You can not sign any additional payments ."
                   
                    End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Function
                ElseIf RsTemp("SumNote_Value").value + val(Me.XPTxtVal.Text) > DblCreditNoteValue Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "⁄›Ê« ..."
                    Msg = Msg & CHR(13) & "·ﬁœ  „  ”ÃÌ· „œ›Ê⁄«  √Ê (⁄„· Œ’Ê„«  „ﬂ ”»…) „”»ﬁ« ·Â–Â «·›« Ê—…"
                    Msg = Msg & CHR(13) & "Ê»≈÷«›… «·ﬁÌ„… «·Õ«·Ì… ”Ê›   ŒÿÏ «·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—…"
                    Msg = Msg & CHR(13) & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„”Ã·…...."
                    Msg = Msg & CHR(13) & "„·ÕÊŸ…:-"
                    Msg = Msg & CHR(13) & "«·ﬁÌ„… «·√Ã·… „‰ «·›« Ê—… ÂÏ : " & DblCreditNoteValue
                    Msg = Msg & CHR(13) & "ﬁÌ„… «·„œ›Ê⁄«  «Ê «·Œ’Ê„«  «·„ﬂ ”»… «·”«»ﬁ… ·Â–Â «·›« Ê—… : " & RsTemp("SumNote_Value").value
                    Else
                    Msg = "Sorry ..."
                    Msg = Msg & CHR(13) & "You have been registered  or work earned discounts in advance for this bill"
                    Msg = Msg & CHR(13) & "and in addition, the present value would exceed the value of these futures bill"
                    Msg = Msg & CHR(13) & "Please review the recorded value...."
                    Msg = Msg & CHR(13) & "Note:-"
                    Msg = Msg & CHR(13) & "Futures value of the invoice is : " & DblCreditNoteValue
                    Msg = Msg & CHR(13) & "The value of payments or discounts earned previous to this bill : " & RsTemp("SumNote_Value").value
                   
                    End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Function
                End If
            End If

        Else
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« «·›« Ê—… —ﬁ„ " & Trim(Me.TxtTransSerial.Text)
            Msg = Msg & CHR(13) & "·Ì”  „”Ã·… „⁄ «·⁄„Ì· " & Me.DBCboClientName.Text
            Else
             Msg = "Sorry invoice number " & Trim(Me.TxtTransSerial.Text)
            Msg = Msg & CHR(13) & "It is not registered with the customer " & Me.DBCboClientName.Text
            
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtTransSerial.SetFocus
            Exit Function
        End If
    End If

    CheckDebitTrans = True
    Exit Function
ErrTrap:
End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "NoteID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_Trans()
    Dim Msg As String
    On Error GoTo ErrTrap

    If XPTxtID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial1.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        Msg = "Confirm Delete"
        End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
    Deletepost Me.Name, "Notes", "NoteID", 0, val(Dcbranch.BoundText), val(XPTxtID.Text), TxtNoteSerial1.Text
    
                rs.delete
                Dim StrSQL As String
                StrSQL = "Delete From notes  Where NoteId=" & val(XPTxtID.Text)
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

                WriteInfo
            End If
        End If

    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This process is not available there are not records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
     Msg = "Sorry...error during delete " & CHR(13)
   
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Function print_Cheque(Optional ChqueNum As String = "", Optional report_no As String = "", Optional serial As String)
    hide_logo = True
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From Expanses_Order  where ChqueNum='0'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
    Else
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    '
    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If RsData.BOF Or RsData.EOF Then
    'GetMsgs 138, vbExclamation
    '    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    RsData.Close
    '    Set RsData = Nothing
    '    Screen.MousePointer = vbDefault
    '    Exit Function
    'End If
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo
    'MsgBox ToHijriDate(Date)

    xReport.ParameterFields(5).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 1, 2)
    xReport.ParameterFields(6).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 4, 2)
    xReport.ParameterFields(7).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 9, 2)

    xReport.ParameterFields(8).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 1, 2)
    xReport.ParameterFields(9).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 4, 2)
    xReport.ParameterFields(10).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 9, 2)
    xReport.ParameterFields(11).AddCurrentValue CStr(XPMTxtRemarks.Text)
    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtVal.Text)
    xReport.ParameterFields(13).AddCurrentValue CStr(Me.txtperson.Text)
    xReport.ParameterFields(14).AddCurrentValue CStr(lbl(18).Caption)
    xReport.ParameterFields(15).AddCurrentValue Format$(DtpChequeDueDate.value, "dd/mm/yyyy")
 
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·”‰œ " & TxtNoteSerial1.Text & CHR(13) & "   «· «—ÌŒ " & XPDtbTrans & CHR(13) & "   ‰Ê⁄ «·„œ›Ê⁄«  " & DCboCashType & CHR(13) & "   «·›—⁄  " & Dcbranch & CHR(13) & "   «·«”„  " & DBCboClientName & CHR(13) & "   ﬁÌ„Â «·„œ›Ê⁄«   " & XPTxtVal & CHR(13) & "   ÿ—Ìﬁ… «·œ›⁄ " & CboPaymentType & CHR(13) & "   «·Œ“Ì‰…  " & DcboBox & CHR(13) & "   «·»‰ﬂ  " & DcboBankName & CHR(13) & "   —ﬁ„ «·‘Ìﬂ  " & TxtChequeNumber & CHR(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & CHR(13) & "     »‰«¡ ⁄·Ï   " & XPMTxtRemarks & CHR(13) & "   «·‘—Õ «·⁄«„    " & txt_general_des & CHR(13) & "   —ﬁ„ «·ﬁÌœ   " & TxtNoteSerial & CHR(13) & "ÿ—› „œÌ‰  " & DcboDebitSide & CHR(13) & " ÿ—› œ«∆‰ " & DcboCreditSide & CHR(13) & "«”„ «·„” Œœ„ " & DCboUserName
                        
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Vchr. NO.  " & TxtNoteSerial1.Text & CHR(13) & "   Date " & XPDtbTrans & CHR(13) & "  Payment Type " & DCboCashType & CHR(13) & "   Branch  " & Dcbranch & CHR(13) & "   Name  " & DBCboClientName & CHR(13) & "  Value" & XPTxtVal & CHR(13) & "   Cash/   Cheque " & CboPaymentType & CHR(13) & "   Box  " & DcboBox & CHR(13) & "   Bank  " & DcboBankName & CHR(13) & "   Cheque No" & TxtChequeNumber & CHR(13) & "  Due Date  " & DtpChequeDueDate & CHR(13) & "  Based On " & XPMTxtRemarks & CHR(13) & "  General Des  " & txt_general_des & CHR(13) & " Ge NO.  " & TxtNoteSerial & CHR(13) & "Debit " & DcboDebitSide & CHR(13) & "Credit " & DcboCreditSide & CHR(13) & " UserName " & DCboUserName
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 50, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 50, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtNoteSerial, TxtNoteSerial1
    End If
    
End Function

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

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hwnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "«·„œ›Ê⁄« ", 1, 15204351, -2147483630
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

Private Sub XPDtbTrans_Change()

    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
End Sub

Private Sub XPTxtVal_Change()
    Me.lbl(18).Caption = WriteNo(Me.XPTxtVal.Text, 0, True)

    If TxtModFlg.Text = "N" Then
        txtAdv_payment_value.Text = XPTxtVal.Text
    End If

End Sub
 
Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.Text, 0)
End Sub

Private Sub WriteInfo()
Exit Sub
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StartWeekDate As Date
    Dim EndWeekDate As Date
    Dim StrTemp As String
    Dim I As Integer

    StartWeekDate = GetWeekStartEND(Date, 0)
    EndWeekDate = DateAdd("d", 7, StartWeekDate)
    StrTemp = "«·≈”»Ê⁄ «·Õ«·Ï „‰ " & DisplayDate(StartWeekDate)
    StrTemp = StrTemp & " ≈·Ï " & DisplayDate(EndWeekDate)
    Me.lbl(22).Caption = StrTemp

    For I = LblLinkInfo.LBound To LblLinkInfo.UBound
        LblLinkInfo(I).Caption = "0"
    Next I

    '------------------------------------------------------------------------------
    '„œ›Ê⁄«  «·ÌÊ„
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 5) "
    StrSQL = StrSQL + " AND NoteDate=" & SQLDate(Date, True)
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For I = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(0).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(1).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(6).Caption = val(Me.LblLinkInfo(0).Caption) + val(Me.LblLinkInfo(1).Caption)
    Else
        Me.LblLinkInfo(0).Caption = 0
        Me.LblLinkInfo(1).Caption = 0
        Me.LblLinkInfo(6).Caption = 0
    End If

    '------------------------------------------------------------------------------
    '„œ›Ê⁄«  «·√”»Ê⁄ «·Õ«·Ï
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 5) "
    StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(StartWeekDate, True)
    StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(EndWeekDate, True)
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For I = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(2).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(3).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(7).Caption = val(Me.LblLinkInfo(2).Caption) + val(Me.LblLinkInfo(3).Caption)
    Else
        Me.LblLinkInfo(0).Caption = 0
        Me.LblLinkInfo(1).Caption = 0
        Me.LblLinkInfo(7).Caption = 0
    End If

    '------------------------------------------------------------------------------
    '„œ›Ê⁄«  «·‘Â— «·Õ«·Ï
    StrSQL = " SELECT     SUM(Note_Value) AS SumX, NoteCashingType"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " Where (NoteType = 5) "
    StrSQL = StrSQL + " AND Month(NoteDate)=" & Month(Date) & ""
    StrSQL = StrSQL + " AND Year(NoteDate)=" & year(Date) & ""
    StrSQL = StrSQL + " GROUP BY NoteCashingType"
    StrSQL = StrSQL + " Order BY NoteCashingType"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For I = 0 To rs.RecordCount - 1

            If rs("NoteCashingType").value = 0 Then
                Me.LblLinkInfo(4).Caption = rs("SumX").value
            ElseIf rs("NoteCashingType").value = 1 Then
                Me.LblLinkInfo(5).Caption = rs("SumX").value
            End If

            rs.MoveNext
        Next

        Me.LblLinkInfo(8).Caption = val(Me.LblLinkInfo(4).Caption) + val(Me.LblLinkInfo(5).Caption)
    Else
        Me.LblLinkInfo(4).Caption = 0
        Me.LblLinkInfo(5).Caption = 0
        Me.LblLinkInfo(8).Caption = 0
    End If

End Sub

