VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReceiptPart 
   BackColor       =   &H00E2E9E9&
   Caption         =   " «· Õ’Ì·«  "
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18090
   HelpContextID   =   430
   Icon            =   "FrmReceiptPart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9885
   ScaleWidth      =   18090
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9885
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18090
      _cx             =   31909
      _cy             =   17436
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
      Align           =   5
      AutoSizeChildren=   7
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
      Begin C1SizerLibCtl.C1Elastic l 
         Height          =   8295
         Index           =   5
         Left            =   10320
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   675
         Width           =   7650
         _cx             =   13494
         _cy             =   14631
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
         Begin VB.ComboBox CboPaymentType 
            Height          =   315
            Left            =   3150
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   4200
            Width           =   2955
         End
         Begin VB.TextBox TxtCustCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   1830
            Width           =   975
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   1200
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   0
            Width           =   1575
         End
         Begin VB.TextBox Txt_akchen 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   -210
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.ComboBox CboType 
            Height          =   315
            Left            =   3180
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1080
            Width           =   2955
         End
         Begin VB.ComboBox CboDealerType 
            Height          =   315
            Left            =   3180
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1470
            Width           =   2955
         End
         Begin VB.TextBox TxtID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   675
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   390
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   840
            Visible         =   0   'False
            Width           =   270
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   1470
            TabIndex        =   19
            Top             =   1830
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker DtbBill 
            Height          =   345
            Left            =   4515
            TabIndex        =   20
            Top             =   375
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   609
            _Version        =   393216
            Format          =   95944705
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   1440
            TabIndex        =   50
            Top             =   720
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   3180
            Index           =   1
            Left            =   240
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   6840
            Width           =   7395
            _cx             =   13044
            _cy             =   5609
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
            Begin VSFlex8Ctl.VSFlexGrid FgDetails 
               Height          =   1245
               Left            =   240
               TabIndex        =   56
               Top             =   195
               Width           =   7140
               _cx             =   12594
               _cy             =   2196
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
               Rows            =   2
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   280
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmReceiptPart.frx":038A
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»«·€ «·„Õ’·… ›Ï «·Õ—ﬂ… «·Õ«·Ì…"
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   8
               Left            =   3330
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   0
               Width           =   4050
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   2325
            Index           =   5
            Left            =   1440
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   4560
            Width           =   6195
            _cx             =   10927
            _cy             =   4101
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
            ForeColor       =   192
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
            Begin VB.TextBox TXTBankName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   450
               Visible         =   0   'False
               Width           =   4845
            End
            Begin VB.TextBox TxtChequeNumber 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2190
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   780
               Width           =   2685
            End
            Begin MSComCtl2.DTPicker DtpChequeDueDate 
               Height          =   315
               Left            =   2190
               TabIndex        =   86
               Top             =   1110
               Width           =   2685
               _ExtentX        =   4736
               _ExtentY        =   556
               _Version        =   393216
               Format          =   95944705
               CurrentDate     =   39614
            End
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   315
               Left            =   30
               TabIndex        =   87
               Top             =   120
               Width           =   4845
               _ExtentX        =   8546
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcChequeBox 
               Height          =   315
               Left            =   30
               TabIndex        =   88
               Top             =   1530
               Width           =   4845
               _ExtentX        =   8546
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCAccounts 
               Height          =   315
               Left            =   60
               TabIndex        =   89
               Top             =   1890
               Width           =   4815
               _ExtentX        =   8493
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBankName 
               Height          =   315
               Left            =   30
               TabIndex        =   96
               Top             =   480
               Width           =   4845
               _ExtentX        =   8546
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Œ“‰…"
               Height          =   285
               Index           =   9
               Left            =   4830
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   150
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«›Ÿ… «·‘Ìﬂ« "
               Height          =   285
               Index           =   43
               Left            =   4800
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   1650
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·»‰ﬂ"
               Height          =   285
               Index           =   23
               Left            =   4830
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·‘Ìﬂ"
               Height          =   285
               Index           =   26
               Left            =   4800
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
               Height          =   285
               Index           =   27
               Left            =   4860
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   1110
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ”«»"
               Height          =   285
               Index           =   31
               Left            =   4800
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   1890
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   135
               Index           =   38
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   150
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   180
               Index           =   35
               Left            =   11040
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   360
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·ﬁ”ÿ"
               Height          =   180
               Index           =   34
               Left            =   11160
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   150
               Width           =   960
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   1995
            Index           =   8
            Left            =   3000
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   2160
            Width           =   4545
            _cx             =   8017
            _cy             =   3519
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
            ForeColor       =   192
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
            Begin VB.TextBox Txt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   345
               Index           =   3
               Left            =   120
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   480
               Visible         =   0   'False
               Width           =   1050
            End
            Begin VB.TextBox TxtValue 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   315
               Left            =   150
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   1515
               Width           =   2955
            End
            Begin VB.TextBox TxtQastNO 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   150
               Locked          =   -1  'True
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   840
               Width           =   2955
            End
            Begin VB.TextBox TxtSum 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   150
               Locked          =   -1  'True
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   1185
               Width           =   2955
            End
            Begin VB.ComboBox CboResType 
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   102
               Top             =   120
               Width           =   2955
            End
            Begin VB.ComboBox CboPrecenType 
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   1560
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   101
               Top             =   480
               Visible         =   0   'False
               Width           =   1530
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„…"
               Height          =   300
               Index           =   17
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   480
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»·€ «·„Õ’·"
               Height          =   300
               Index           =   6
               Left            =   3450
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   1530
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·√ﬁ”«ÿ"
               Height          =   300
               Index           =   24
               Left            =   3330
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   870
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Ã„Ê⁄  «· Õ’Ì·«  "
               Height          =   300
               Index           =   25
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   1200
               Width           =   1320
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «· Õ’Ì·"
               Height          =   300
               Index           =   3
               Left            =   3330
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   180
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·Œ’„"
               Height          =   300
               Index           =   16
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   480
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·ﬁ”ÿ"
               Height          =   180
               Index           =   45
               Left            =   11160
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   150
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   180
               Index           =   44
               Left            =   11040
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   360
               Width           =   960
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   135
               Index           =   42
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   150
               Visible         =   0   'False
               Width           =   945
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·ﬁ»÷"
            Height          =   315
            Index           =   21
            Left            =   6210
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   4320
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   315
            Index           =   20
            Left            =   6720
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   720
            Width           =   675
         End
         Begin VB.Label Lb_note_value_by_characters 
            Alignment       =   1  'Right Justify
            Height          =   495
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   8280
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   270
            Index           =   18
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   60
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„·Ì…"
            Height          =   270
            Index           =   11
            Left            =   6420
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1140
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„Ì·"
            Height          =   270
            Index           =   10
            Left            =   6420
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1500
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   270
            Index           =   5
            Left            =   6420
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   60
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· Õ’Ì·"
            Height          =   270
            Index           =   0
            Left            =   6420
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   405
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   270
            Index           =   7
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1815
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8295
         Index           =   2
         Left            =   30
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   675
         Width           =   10275
         _cx             =   18124
         _cy             =   14631
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
            Height          =   2310
            Index           =   7
            Left            =   75
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   5970
            Width           =   10110
            _cx             =   17833
            _cy             =   4075
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
            ForeColor       =   192
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
            Align           =   0
            AutoSizeChildren=   7
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
            Begin VB.TextBox TXTMessageDES 
               Alignment       =   1  'Right Justify
               Height          =   225
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   180
               Width           =   2730
            End
            Begin VB.TextBox TxtRemark 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   525
               Left            =   2970
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   59
               Top             =   450
               Width           =   4965
            End
            Begin MSDataListLib.DataCombo DcboDebitSide 
               Height          =   315
               Left            =   45
               TabIndex        =   36
               Top             =   360
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCreditSide 
               Height          =   315
               Left            =   45
               TabIndex        =   37
               Top             =   870
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboDebitSide1 
               Height          =   315
               Left            =   45
               TabIndex        =   41
               Top             =   615
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCreditSide1 
               Height          =   315
               Left            =   45
               TabIndex        =   42
               Top             =   1125
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1050
               Index           =   9
               Left            =   3120
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   1080
               Width           =   6795
               _cx             =   11986
               _cy             =   1852
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
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
               ForeColor       =   192
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
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   4080
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   240
                  Width           =   2055
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   255
                  Index           =   9
                  Left            =   2520
                  TabIndex        =   118
                  Top             =   240
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   450
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   255
                  Index           =   11
                  Left            =   840
                  TabIndex        =   119
                  Top             =   240
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   450
                  ButtonPositionImage=   1
                  Caption         =   "ÿ»«⁄… ”‰œ «·ﬁ»÷"
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
               Begin MSComCtl2.DTPicker DTPickerAccFrom 
                  Height          =   255
                  Left            =   4080
                  TabIndex        =   120
                  Top             =   600
                  Width           =   2100
                  _ExtentX        =   3704
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   95944705
                  CurrentDate     =   38784
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   255
                  Index           =   12
                  Left            =   840
                  TabIndex        =   121
                  Top             =   600
                  Width           =   3135
                  _ExtentX        =   5530
                  _ExtentY        =   450
                  ButtonPositionImage=   1
                  Caption         =   " ÿ»«⁄Â ﬂ‘› Õ”«» «·⁄„Ì·"
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Height          =   60
                  Index           =   48
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   75
                  Visible         =   0   'False
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   90
                  Index           =   47
                  Left            =   15630
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   165
                  Width           =   1365
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  «·ﬁ”ÿ"
                  Height          =   75
                  Index           =   46
                  Left            =   15810
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   75
                  Width           =   1350
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   600
               Index           =   10
               Left            =   120
               TabIndex        =   122
               TabStop         =   0   'False
               Top             =   1560
               Width           =   2895
               _cx             =   5106
               _cy             =   1058
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
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
               ForeColor       =   192
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "œ·«·«  «·«·Ê«‰"
               Align           =   0
               AutoSizeChildren=   7
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
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "«·ﬁ”ÿ „”œœ Ã“¡ „‰…"
                  Height          =   255
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.Shape Shape1 
                  FillColor       =   &H0000C000&
                  FillStyle       =   0  'Solid
                  Height          =   255
                  Left            =   1680
                  Top             =   240
                  Width           =   255
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»Ì«‰«  «·ﬁ”ÿ"
                  Height          =   45
                  Index           =   36
                  Left            =   6720
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   45
                  Width           =   600
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   60
                  Index           =   32
                  Left            =   6675
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   90
                  Width           =   570
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Height          =   30
                  Index           =   19
                  Left            =   45
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   585
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  «·ﬁ”ÿ"
               Height          =   225
               Index           =   28
               Left            =   7875
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   180
               Width           =   675
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   240
               Index           =   22
               Left            =   8040
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   570
               Width           =   660
            End
            Begin VB.Label lbl3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› œ«∆‰2"
               Height          =   240
               Index           =   17
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   1125
               Width           =   660
            End
            Begin VB.Label lbl3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› „œÌ‰2 "
               Height          =   240
               Index           =   16
               Left            =   1935
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   615
               Width           =   825
            End
            Begin VB.Label lbl3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› „œÌ‰"
               Height          =   240
               Index           =   32
               Left            =   1935
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   360
               Width           =   825
            End
            Begin VB.Label lbl3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› œ«∆‰"
               Height          =   240
               Index           =   31
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   870
               Width           =   660
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   ".."
               ForeColor       =   &H000000FF&
               Height          =   270
               Index           =   30
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   1365
               Width           =   2520
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·› —… :"
               Height          =   180
               Index           =   29
               Left            =   735
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   180
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label LblDevID 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   180
               Left            =   1575
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   180
               Index           =   33
               Left            =   60
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   180
               Visible         =   0   'False
               Width           =   660
            End
         End
         Begin C1SizerLibCtl.C1Tab TabMain 
            Height          =   5925
            Left            =   0
            TabIndex        =   7
            Top             =   30
            Width           =   10260
            _cx             =   18098
            _cy             =   10451
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
            Caption         =   " «· Õ’Ì·«  «·„ƒÃ·…| «· Õ’Ì·«  ”«»ﬁ…"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   3
            Position        =   0
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
            Picture(0)      =   "FrmReceiptPart.frx":0525
            Picture(1)      =   "FrmReceiptPart.frx":0ABF
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5460
               Index           =   3
               Left            =   10905
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   420
               Width           =   10170
               _cx             =   17939
               _cy             =   9631
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
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmReceiptPart.frx":0E59
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   2295
                  Index           =   8
                  Left            =   30
                  TabIndex        =   30
                  Top             =   3135
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   4048
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
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   285
                  Index           =   7
                  Left            =   30
                  TabIndex        =   29
                  Top             =   30
                  Width           =   885
                  _ExtentX        =   1561
                  _ExtentY        =   503
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
                  DrawFocusRectangle=   0   'False
               End
               Begin VSFlex8UCtl.VSFlexGrid FgReceipted 
                  Height          =   2790
                  Left            =   30
                  TabIndex        =   9
                  Top             =   330
                  Width           =   10110
                  _cx             =   17833
                  _cy             =   4921
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
                  BackColorBkg    =   16777215
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   10
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmReceiptPart.frx":0EF2
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
               Begin VSFlex8UCtl.VSFlexGrid FgPayed 
                  Height          =   2295
                  Left            =   30
                  TabIndex        =   26
                  Top             =   3135
                  Width           =   10110
                  _cx             =   17833
                  _cy             =   4048
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
                  BackColorBkg    =   16777215
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   10
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmReceiptPart.frx":10A0
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "√ﬁ”«ÿ ”«»ﬁ… ”œœ  ≈·Ï «·⁄„Ì·"
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
                  Height          =   2295
                  Index           =   15
                  Left            =   930
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   3135
                  Width           =   9210
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "√ﬁ”«ÿ ”«»ﬁ… Õ’·  „‰ «·⁄„Ì·"
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
                  Index           =   14
                  Left            =   930
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   30
                  Width           =   9210
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5460
               Index           =   4
               Left            =   45
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   420
               Width           =   10170
               _cx             =   17939
               _cy             =   9631
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
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmReceiptPart.frx":124E
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VSFlex8Ctl.VSFlexGrid Fg 
                  Height          =   4530
                  Left            =   30
                  TabIndex        =   11
                  Top             =   510
                  Width           =   10110
                  _cx             =   17833
                  _cy             =   7990
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
                  Rows            =   2
                  Cols            =   14
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   280
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmReceiptPart.frx":12CF
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
                  WallPaperAlignment=   9
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin VB.Image Img 
                  Height          =   465
                  Left            =   9450
                  Picture         =   "FrmReceiptPart.frx":152E
                  Top             =   30
                  Width           =   690
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   375
                  Index           =   13
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   5055
                  Width           =   10110
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
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
                  Height          =   465
                  Index           =   12
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   30
                  Width           =   9405
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   630
         Index           =   6
         Left            =   75
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   18060
         _cx             =   31856
         _cy             =   1111
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   20.25
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
         Caption         =   " «· Õ’Ì·«  "
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
         Begin VB.TextBox XPTxtID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   120
            Visible         =   0   'False
            Width           =   1440
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   315
            Index           =   0
            Left            =   2895
            TabIndex        =   2
            Top             =   135
            Width           =   1065
            _ExtentX        =   1879
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
            ButtonImage     =   "FrmReceiptPart.frx":18B8
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
            Left            =   1470
            TabIndex        =   3
            Top             =   135
            Width           =   1350
            _ExtentX        =   2381
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
            ButtonImage     =   "FrmReceiptPart.frx":1C52
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
            Left            =   4035
            TabIndex        =   4
            Top             =   135
            Width           =   1155
            _ExtentX        =   2037
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
            ButtonImage     =   "FrmReceiptPart.frx":1FEC
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
            Left            =   165
            TabIndex        =   5
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
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
            ButtonImage     =   "FrmReceiptPart.frx":2386
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
         Height          =   795
         Index           =   0
         Left            =   2790
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   9105
         Width           =   12225
         _cx             =   21564
         _cy             =   1402
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
            Height          =   270
            Index           =   0
            Left            =   7155
            TabIndex        =   63
            Top             =   330
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   476
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
            Height          =   270
            Index           =   1
            Left            =   6240
            TabIndex        =   64
            Top             =   330
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   476
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
            Height          =   270
            Index           =   2
            Left            =   5190
            TabIndex        =   65
            Top             =   345
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   476
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
            Height          =   270
            Index           =   3
            Left            =   4215
            TabIndex        =   66
            Top             =   330
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   476
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
            Height          =   270
            Index           =   4
            Left            =   3165
            TabIndex        =   67
            Top             =   330
            Width           =   840
            _ExtentX        =   1482
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
            Height          =   270
            Index           =   6
            Left            =   75
            TabIndex        =   68
            Top             =   330
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   270
            Left            =   555
            TabIndex        =   69
            Top             =   -330
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   476
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4950
            TabIndex        =   70
            Top             =   60
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   270
            Index           =   5
            Left            =   2130
            TabIndex        =   71
            Top             =   330
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   476
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
            Height          =   255
            Index           =   10
            Left            =   1425
            TabIndex        =   72
            Top             =   345
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   450
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
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   255
            Index           =   13
            Left            =   645
            TabIndex        =   73
            Top             =   345
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   450
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«—”«· —”«·…"
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
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   4
            Left            =   1785
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   2
            Left            =   3765
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   60
            Width           =   1080
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   240
            Left            =   1485
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   60
            Width           =   675
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   240
            Left            =   2835
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   60
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   240
            Index           =   1
            Left            =   7110
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   60
            Width           =   1005
         End
      End
      Begin ImpulseButton.ISButton CmdAttach 
         Height          =   375
         Left            =   1470
         TabIndex        =   79
         Top             =   9480
         Width           =   1170
         _ExtentX        =   2064
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
   End
End
Attribute VB_Name = "FrmReceiptPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RsNotes As ADODB.Recordset
Dim TTP As clstooltip
Dim cDcboSearch(2) As clsDCboSearch
Dim Sanad_No As Integer
Dim notesType As Integer
Dim SngUseValue As Single ' „'ﬁÌ„… «·Œ’
Dim dBox As Integer
    
Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "—ﬁ„ «·”‰œ " & TxtNoteSerial1.Text & CHR(13) & "    «—ÌŒ «· Õ’Ì· " & DtbBill & CHR(13) & "  «·›—⁄  " & Dcbranch & CHR(13) & "     ‰Ê⁄ «·⁄„·Ì…  " & CboType & CHR(13) & "    ‰Ê⁄ «·⁄„Ì· " & CboDealerType & CHR(13) & "   «”„ «·⁄„Ì· " & DBCboClientName & CHR(13) & "  ⁄œœ «·«ﬁ”«ÿ " & TxtQastNO & CHR(13) & "   „Ã„Ê⁄ «·«ﬁ”«ÿ " & TxtSum & CHR(13) & "    «·„»·€ «·„Õ’·" & TxtValue & CHR(13) & " ÿ—Ìﬁ… «·ﬁ»÷   " & CboPayMentType & CHR(13) & "  «”„ «·Œ“Ì‰… " & DcboBox & CHR(13) & "  «”„ «·»‰ﬂ  " & TXTBankName & CHR(13) & "  —ﬁ„ «·‘Ìﬂ  " & TxtChequeNumber & CHR(13) & "    «·«” Õﬁ«ﬁ " & DtpChequeDueDate & CHR(13) & "   Õ«›Ÿ… «·‘Ìﬂ«  " & DCChequeBox
        LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & "Vchr No " & TxtNoteSerial1.Text & CHR(13) & "  Collect Date " & DtbBill & CHR(13) & "  Branch  " & Dcbranch & CHR(13) & "     Operation Type    " & CboType & CHR(13) & "    Customer Type " & CboDealerType & CHR(13) & "  Customer/supplier Name " & DBCboClientName & CHR(13) & "No oF Installments" & TxtQastNO & CHR(13) & "    Installments Total " & TxtSum & CHR(13) & "  Payed " & TxtValue & CHR(13) & "Payment Type    " & CboPayMentType & CHR(13) & "  «Box Name " & DcboBox & CHR(13) & "  Bank Name    " & TXTBankName & CHR(13) & "    Cheque No  " & TxtChequeNumber & CHR(13) & "  Due Date " & DtpChequeDueDate & CHR(13) & "   Cheque Box   " & DCChequeBox
       Dim NoteType As Integer

    If CboType.ListIndex = 0 Then
        NoteType = 18
    ElseIf CboType.ListIndex = 1 Then
        NoteType = 19
    End If
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), NoteType, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg, , , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), NoteType, Date, Time, LogTextA, LogTextE, Me.Name, "D", , , val(TxtNoteSerial), val(TxtNoteSerial1)
    End If
    
End Function
    
Private Sub CboPayMentType_Change()

    If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
        DcboBankName.Text = ""
        TxtChequeNumber.Text = ""
        Me.DcboBox.Text = ""
        DCChequeBox.Text = ""
        TXTBankName.Text = ""
        DCAccounts.Text = ""
    End If

    DCChequeBox.Enabled = False

    If SystemOptions.UserInterface = ArabicInterface Then
        lbl(26).Caption = "—ﬁ„ «·‘Ìﬂ"
        lbl(27).Caption = " «—ÌŒ «·«” Õﬁ«ﬁ"
    
    Else
        lbl(26).Caption = "Cheque No"
        lbl(27).Caption = "Due Date"
    End If
    
    If Me.CboPayMentType.ListIndex = 0 Then
      DCAccounts.Enabled = False
 DCAccounts.Text = ""
        Me.lbl(9).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(23).Enabled = False
        Me.lbl(26).Enabled = False
        Me.lbl(27).Enabled = False
        
             Me.lbl(43).Enabled = False
        Me.lbl(31).Enabled = False
        
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False

        '    Frame3.Enabled = False
        If Me.TxtModFlg.Text <> "R" Then
            GetUserData user_id, , , , dBox
   
            Me.DcboBox.BoundText = dBox
        End If

    ElseIf Me.CboPayMentType.ListIndex = 1 Then
  DCAccounts.Enabled = False
 DCAccounts.Text = ""
        If SystemOptions.ChequeBox = True Then
            TXTBankName.Visible = True
            DCChequeBox.Enabled = True
        Else
            TXTBankName.Visible = False
                     Me.lbl(43).Enabled = False
      
        End If
  Me.lbl(31).Enabled = False
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(23).Enabled = True
        Me.lbl(26).Enabled = True
        Me.lbl(27).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
    
        'Frame3.Enabled = False
    ElseIf Me.CboPayMentType.ListIndex = 2 Then
   DCAccounts.Enabled = False
 DCAccounts.Text = ""
        TXTBankName.Visible = False
 
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(23).Enabled = True
        Me.lbl(26).Enabled = True
        Me.lbl(27).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        TXTBankName.Visible = False

        'Frame3.Enabled = True
        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(26).Caption = "—ﬁ„ «·ÕÊ«·Â"
            lbl(27).Caption = " «—ÌŒÂ«"
    
        Else
            lbl(26).Caption = "Transfer No"
            lbl(27).Caption = "Date"
        End If
           Me.lbl(43).Enabled = False
        Me.lbl(31).Enabled = False
        
    
    ElseIf Me.CboPayMentType.ListIndex = 3 Then
  DCAccounts.Enabled = False
 DCAccounts.Text = ""
        TXTBankName.Visible = False
 
        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(23).Enabled = True
        Me.lbl(26).Enabled = True
        Me.lbl(27).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        TXTBankName.Visible = False
  Me.lbl(31).Enabled = False
        'Frame3.Enabled = True
                        If SystemOptions.UserInterface = ArabicInterface Then
                            lbl(26).Caption = "—ﬁ„ «·‘Ìﬂ"
                            lbl(27).Caption = " «—ÌŒÂ"
                    
                        Else
                            lbl(26).Caption = "Transfer No"
                            lbl(27).Caption = "Date"
                        End If
          Me.lbl(43).Enabled = False
        Me.lbl(31).Enabled = False
        
     ElseIf Me.CboPayMentType.ListIndex = 4 Then

 DCAccounts.Enabled = True
 DcboBankName.Enabled = False
 
         TXTBankName.Visible = False
        Me.DcboBox.Enabled = False
 
        'Me.DcboBankName.Enabled = True
         Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        TXTBankName.Visible = False
    Me.lbl(23).Enabled = False
        Me.lbl(26).Enabled = False
        Me.lbl(27).Enabled = False
               Me.lbl(43).Enabled = False
        
                          
                        If SystemOptions.UserInterface = ArabicInterface Then
                            lbl(31).Caption = "«·Õ”«»"
                             
                    
                        Else
                            lbl(26).Caption = "Account #"
                             
                        End If
                        
                        
    Else

        Me.lbl(9).Enabled = False
        Me.DcboBox.Enabled = False
        Me.lbl(23).Enabled = False
        Me.lbl(26).Enabled = False
        Me.lbl(27).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        DCAccounts.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub CboDealerType_Change()
    Dim Dcombos As ClsDataCombos

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Set Dcombos = New ClsDataCombos

        If Me.CboDealerType.ListIndex = 0 Then
            Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, False
        ElseIf Me.CboDealerType.ListIndex = 1 Then
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, False
        End If

        cDcboSearch(0).Refresh
    End If

    WriteDev
End Sub

Private Sub CboDealerType_Click()
    CboDealerType_Change
End Sub

Private Sub CboPrecenType_Change()
 
    With CboPrecenType

        If .ListIndex > -1 Then

            Select Case .ItemData(.ListIndex)

                Case 1
                    lbl(17).Caption = "«·‰”»… "

                Case 2
                    lbl(17).Caption = "«·ﬁÌ„… "

                Case 3
                    Txt(3).Text = 0
            End Select

            CalPre
        End If

    End With
 
    Exit Sub
End Sub

Private Sub CalPre()
    On Error GoTo ErrTrap

    Dim SngAllValue As Single

    'Õ”«» ﬁÌ„… «·›«∆œ…
    If Me.CboPrecenType.ListIndex > -1 Then
        If Me.CboPrecenType.ItemData(CboPrecenType.ListIndex) = 1 Then
            SngUseValue = (val(TxtSum.Text) * val(Txt(3).Text)) / 100
        ElseIf Me.CboPrecenType.ItemData(CboPrecenType.ListIndex) = 2 Then
            SngUseValue = val(Me.Txt(3).Text)
        ElseIf Me.CboPrecenType.ItemData(CboPrecenType.ListIndex) = 3 Then
            SngUseValue = 0
        End If
    End If

    'TxtValue.text = (SngUseValue)
    '«·„»·€ «·ﬂ·Ï («·–Ï ”Ê› Ìﬁ”ÿ) Ì”«ÊÏ Õ”«» ﬁÌ„…
    '«·›«∆œ… „⁄ ﬁÌ„… «·„»·€ «·„ »ﬁÏ
    SngAllValue = val(TxtSum.Text) - SngUseValue
    TxtValue.Text = (SngAllValue)
 
    Exit Sub
ErrTrap:
End Sub

Private Sub CboResType_Change()
    SetReleaseType
    WriteDev
End Sub

Private Sub CboPrecenType_Click()
    CboPrecenType_Change
End Sub

Private Sub CboResType_Click()
    SetReleaseType
End Sub

Private Sub CboType_Change()
    GetDealerInstallment

    WriteDev
End Sub

Private Sub CboType_Click()
    GetDealerInstallment

    If Me.TxtModFlg <> "R" Then
        CboDealerType.ListIndex = CboType.ListIndex
    End If

    WriteDev
End Sub

Function SendMessage(currentOpt As Integer)
    Dim Opt As Integer
    Dim CurrentMessage As String
    CurrentMessage = ComposMessage(Me.Name, 0, "", TXTMessageDES.Text, Opt)

    If Opt = currentOpt Then
        SMSSeTTings.SendMessage CurrentMessage, GetCustomerNumber(DBCboClientName.BoundText)
        SMSSeTTings.Hide
    End If
 
End Function

Private Sub Cmd_Click(Index As Integer)
     On Error GoTo ErrTrap
    Dim Reports As ClsRepoerts
    Dim StrSQL As String
    Dim Msg As String
    Dim Opt As Integer
    Dim t As String
    Dim CurrentMessage As String

    Select Case Index

        Case 0
            CboResType.Enabled = True

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            ClearMe
            TxtModFlg.Text = "N"
            Me.DCboUserName.BoundText = user_id
            Me.Dcbranch.BoundText = Current_branch
            CboType.ListIndex = 0
    
            CboResType.ListIndex = 0
           
        Case 1
            
             If ChekClodePeriod(DtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            CboResType.Enabled = False

            If CboResType.ListIndex = 3 Or CboResType.ListIndex = 4 Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Modify Adanced Payments"
                Else
                    Msg = "  ›Ì Õ«·… «·œ›⁄Â «·„ﬁœ„… «Ê œ›⁄Â „‰ «·Õ”«»  ·« Ì„ﬂ‰ «· ⁄œÌ· "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 
                Screen.MousePointer = vbDefault
                Exit Sub
    
            End If
    
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
        
            If SystemOptions.ChequeBox = True And CboPayMentType.ListIndex = 1 Then
         
                If ChequeBoxOperations(val(Me.XPTxtID)) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If
            
            TxtModFlg.Text = "E"
    '        DCAccounts_Change
    '        DBCboClientName_Change
    '        DcChequeBox_Change
    '        DcboBox_Change
    '
    '        DcboBankName_Click (0)
            
            Me.DCboUserName.BoundText = user_id
            CuurentLogdata

        Case 2
         
             If ChekClodePeriod(DtbBill.value) = True Then
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

            '       If Me.TxtModFlg.text = "N" Then
    
            If CboType.ListIndex = 0 Then
                Sanad_No = 25
                notesType = 18
            ElseIf CboType.ListIndex = 1 Then
                Sanad_No = 26
                notesType = 19
            Else
                Msg = "ÌÃ»  ÕœÌœ  ‰Ê⁄ «·⁄„Ì·… Â· (  Õ’Ì· √ﬁ”«ÿ «„  ”œÌœ √ﬁ”«ÿ)..!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                CboType.SetFocus
                SendKeys "{F4}"
                Exit Sub
    
            End If
       
            SaveData
            'SendMessage (1)
     
        Case 3
            Undo

        Case 4
             If ChekClodePeriod(DtbBill.value) = True Then
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
        
            If SystemOptions.ChequeBox = True And CboPayMentType.ListIndex = 1 Then
         
                If ChequeBoxOperations(val(Me.XPTxtID)) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If
            
            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmReceiptPartSearch.show vbModal

        Case 6
            Unload Me

        Case 9
        
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc Me.TxtNoteSerial.Text, , 200
          
        Case 10
          
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            
            End If

            Set Reports = New ClsRepoerts
                
            StrSQL = "SELECT     dbo.Qest_Had_Receipted.QestID, dbo.Qest_Had_Receipted.QeqtNum, dbo.Qest_Had_Receipted.[Value], dbo.Qest_Had_Receipted.Transaction_ID, "
            StrSQL = StrSQL & " dbo.Qest_Had_Receipted.CustID, dbo.Qest_Had_Receipted.ReceiptID, dbo.Qest_Had_Receipted.Status, dbo.Qest_Had_Receipted.ReceiptDate,"
            StrSQL = StrSQL & "dbo.Qest_Had_Receipted.CusName, dbo.Qest_Had_Receipted.Transaction_Serial, dbo.Qest_Had_Receipted.Type, dbo.Qest_Had_Receipted.ReceiptType,"
            StrSQL = StrSQL & "dbo.Qest_Had_Receipted.TransactionTypeName, dbo.Qest_Had_Receipted.QestValue, dbo.Transactions.NoteSerial1 AS BillNoteSerial,"
            StrSQL = StrSQL & " dbo.ReceiptQest.noteserial1"
            StrSQL = StrSQL & "  FROM         dbo.Qest_Had_Receipted INNER JOIN"
            StrSQL = StrSQL & "  dbo.ReceiptQest ON dbo.Qest_Had_Receipted.ReceiptID = dbo.ReceiptQest.ReceiptID LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.Transactions ON dbo.Qest_Had_Receipted.Transaction_ID = dbo.Transactions.Transaction_ID"
            StrSQL = StrSQL & "  where  dbo.Qest_Had_Receipted.ReceiptID =" & val(txtid.Text)
            Reports.ShowSallingTime StrSQL, , , , 19, , ""
            SendMessage (2)

        Case 7
           
            Set Reports = New ClsRepoerts
                
            StrSQL = " SELECT     TOP 100 PERCENT dbo.Transactions.NoteSerial1 AS billNoteSerial1, dbo.Qest_Had_Receipted.QestID, dbo.Qest_Had_Receipted.QeqtNum, "
            StrSQL = StrSQL & "  dbo.Qest_Had_Receipted.[Value], dbo.Qest_Had_Receipted.Transaction_ID, dbo.Qest_Had_Receipted.CustID, dbo.Qest_Had_Receipted.ReceiptID,"
            StrSQL = StrSQL & " dbo.Qest_Had_Receipted.ReceiptDate, dbo.Qest_Had_Receipted.Status, dbo.Qest_Had_Receipted.Transaction_Serial, dbo.Qest_Had_Receipted.CusName,"
            StrSQL = StrSQL & "  dbo.Qest_Had_Receipted.Type, dbo.Qest_Had_Receipted.TransactionTypeName, dbo.Qest_Had_Receipted.QestValue, dbo.Qest_Had_Receipted.ReceiptType,"
            StrSQL = StrSQL & " dbo.ReceiptQest.noteserial1"
            StrSQL = StrSQL & " FROM         dbo.Qest_Had_Receipted INNER JOIN"
            StrSQL = StrSQL & " dbo.ReceiptQest ON dbo.Qest_Had_Receipted.ReceiptID = dbo.ReceiptQest.ReceiptID LEFT OUTER JOIN"
            StrSQL = StrSQL & " dbo.Transactions ON dbo.Qest_Had_Receipted.Transaction_ID = dbo.Transactions.Transaction_ID"
            StrSQL = StrSQL & " Where (dbo.Qest_Had_Receipted.custid = " & val(DBCboClientName.BoundText) & ") And (dbo.Qest_Had_Receipted.type = 0)"
            StrSQL = StrSQL & "  ORDER BY dbo.Qest_Had_Receipted.ReceiptID, dbo.Qest_Had_Receipted.QestID"

            Reports.ShowSallingTime StrSQL, , , , 20, , ""

        Case 12
 
            updateopeningbalanceNewFromsql DTPickerAccFrom.value, Date, True, 0, 0, DcboCreditSide.BoundText, 3

            ShowReport DcboCreditSide.BoundText, DcboCreditSide.Text, DTPickerAccFrom.value, Date

        Case 11

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
        
            If TxtNoteSerial <> "" Then
                If CboPayMentType.ListIndex = 1 Or CboPayMentType.ListIndex = 3 Then    '‘Ìﬂ
                    print_report TxtNoteSerial, Me.TxtNoteSerial1.Text, TXTBankName.Text, CboPayMentType.Text, DcboBox.Text, TxtCustCode.Text, TxtValue.Text, DtbBill.value, txtRemark.Text
                Else 'ÕÊ«·Â
                    print_report TxtNoteSerial, Me.TxtNoteSerial1.Text, DcboBankName.Text, CboPayMentType.Text, DcboBox.Text, TxtCustCode.Text, TxtValue.Text, DtbBill.value, txtRemark.Text
                End If
            End If

        Case 13
            SendMessage (0)
    End Select

    Exit Sub
ErrTrap:
End Sub

Function print_report(Optional NoteSerial As String, Optional NoteSerial1 As String, Optional BankName As String, Optional PaymentType As String, Optional Box As String, Optional Custcode As String, Optional sValue As Double, Optional NoteDate As Date, Optional Remark As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From payment_voucher  where noteserial='" & NoteSerial & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "Payment_voucher1.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Payment_voucher1.rpt"
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
        '    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '    RsData.Close
        '    Set RsData = Nothing
        '    Screen.MousePointer = vbDefault
        '    Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.Text
   
        StrReportTitle = "" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        xReport.ParameterFields(5).AddCurrentValue DcboCreditSide.Text ' «·œ«∆‰
        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    '
    xReport.ParameterFields(6).AddCurrentValue NoteSerial1

    xReport.ParameterFields(7).AddCurrentValue BankName
    xReport.ParameterFields(8).AddCurrentValue PaymentType
    xReport.ParameterFields(9).AddCurrentValue Box
    xReport.ParameterFields(10).AddCurrentValue Custcode

    xReport.ParameterFields(11).AddCurrentValue CStr(sValue)
    xReport.ParameterFields(12).AddCurrentValue WriteNo(CStr(sValue), 0, True)
    xReport.ParameterFields(13).AddCurrentValue CStr(NoteDate)
    xReport.ParameterFields(14).AddCurrentValue CStr(Remark)
    xReport.ParameterFields(15).AddCurrentValue ToHijriDate(NoteDate)

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

Private Sub Cmd_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    Select Case Index

        Case 13:
            Dim Opt As Integer
            Dim CurrentMessage As String
            Cmd(13).ToolTipText = ComposMessage(Me.Name, 0, "", Me.TXTMessageDES.Text, Opt)
    End Select

End Sub

Private Sub CmdAttach_Click()
     On Error Resume Next
ShowAttachments TxtNoteSerial1, "0812201401"

End Sub

Private Sub DataCombo3_Click(Area As Integer)
End Sub

Private Sub DBCboClientName_Change()
    TxtCustCode.Text = ""

    Dim DefaultSalesPersonId As Integer
    Dim Fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, Fullcode

    TxtCustCode.Text = Fullcode

    GetDealerInstallment
    WriteDev
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If CboDealerType.ListIndex = 0 Then
        If KeyCode = vbKeyF3 Then
            FrmCustemerSearch.SearchType = 5
            FrmCustemerSearch.show vbModal
            
        End If

    ElseIf CboDealerType.ListIndex = 1 Then

        If KeyCode = vbKeyF3 Then
            FrmCompanySearch.lblSearchtype.Caption = 6
            FrmCompanySearch.show vbModal
                 
        End If

    End If

End Sub

Private Sub DcboBankName_Click(Area As Integer)

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String
    Dim Account_Code_dynamic As String

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        'Me.DcboDebitSide.BoundText =   "a1a2a4"
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If SystemOptions.ChequeBox = True Then
            Me.DcboDebitSide.BoundText = ""
        Else

            If SystemOptions.banks_Accounts3 = True Then
                Me.DcboDebitSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code1")
            Else
                Me.DcboDebitSide.BoundText = RsSavRec.Fields("Account_Code").value
                     
            End If
        End If

        If CboPayMentType.ListIndex = 2 Or CboPayMentType.ListIndex = 3 Then
                     
            Me.DcboDebitSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If

    End If

End Sub

Private Sub DcboBox_Change()
    WriteDev
End Sub

Private Sub Dcbranch_Click(Area As Integer)

    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
End Sub

Private Sub DcChequeBox_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCodeRefined("TblBoxesData", "BoxID", val(Me.DCChequeBox.BoundText), "Account_Code1")
    End If

End Sub

Private Sub DtbBill_Change()

    If Me.TxtModFlg.Text = "E" Then
 
        CuurentLogdata ("D")
    End If
 
    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""

End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    On Error GoTo ErrTrap

    Dim RowNum As Integer
    Dim IntCounter As Integer
    Dim DblSum As Double
    Dim Payment As Double
    Dim NoteType As Integer
    Dim CurrentMonth As String
    CurrentMonth = ""

    If CboType.ListIndex = 0 Then
        NoteType = 18
    ElseIf CboType.ListIndex = 1 Then
        NoteType = 19
    End If

    Dim Msg As String
    '›Ì Õ«·… œ›⁄Â „‰ «·ﬁ”ÿ ·«»œ „‰ «Œ Ì«— ﬁ”ÿ Ê«Õœ ›ﬁÿ
    Dim NoOFReleased As Integer
    NoOFReleased = 0

    With FG

        For RowNum = .FixedRows To .Rows - 1

            If .Cell(flexcpChecked, RowNum, .ColIndex("Released")) = flexChecked Then
                NoOFReleased = NoOFReleased + 1

                If NoOFReleased > 1 Then
                    Msg = "⁄‰œ  Õ’Ì· œ›⁄…  Õ  Õ”«» «·ﬁ”ÿ " & CHR(13)
                    Msg = Msg & "·«»œ Ê«‰ ÌﬂÊ‰ ﬁ”ÿ Ê«Õœ ›ﬁÿ " & CHR(13)
                    Msg = Msg & "»—Ã«¡ „—«Ã⁄… «·√ﬁ”«ÿ «·„Õ’·… «Ê  €Ì— " & CHR(13)
                    Msg = Msg & "‰Ê⁄ «· Õ’Ì·...!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    
                    .Cell(flexcpChecked, RowNum, .ColIndex("Released")) = flexUnchecked
        
                    Exit Sub
                End If
        
            End If

        Next

    End With

    With FG

        For RowNum = .FixedRows To .Rows - 1

            If .Cell(flexcpChecked, RowNum, .ColIndex("Released")) = flexChecked Then
                IntCounter = IntCounter + 1
                DblSum = DblSum + val(.TextMatrix(RowNum, .ColIndex("Value")))

                If IsNumeric(val(.TextMatrix(RowNum, .ColIndex("Des")))) Then
                    Payment = Payment + val(.TextMatrix(RowNum, .ColIndex("Des")))
                End If
            
                If CboResType.ListIndex = 2 Then
                    CurrentMonth = CurrentMonth & MonthName(Month(.TextMatrix(RowNum, .ColIndex("Due_Date")))) & " Ã“¡ ›ﬁÿ"
                Else
                    CurrentMonth = CurrentMonth & " Ê" & MonthName(Month(.TextMatrix(RowNum, .ColIndex("Due_Date"))))
                End If
      
                '            LstQasts.AddItem Write_Qast(Val(.TextMatrix(RowNum, .ColIndex("Note_serial"))))
                '            LstQasts.ItemData(LstQasts.ListCount - 1) = RowNum
                '            LstNoteID.AddItem .TextMatrix(RowNum, .ColIndex("Note_ID"))
          
            End If
    
        Next RowNum

        If CboResType.ListIndex = 2 Then
        Else
            CurrentMonth = mId(CurrentMonth, 3, Len(CurrentMonth) - 1)
        End If
   
        Me.TXTMessageDES.Text = CurrentMonth
          
        '''''''''''''
 
        '           If Me.TxtModFlg = "E" Then
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
        If .Cell(flexcpChecked, Row, .ColIndex("Released")) = flexChecked Then
            LogTextA = "   ÕœÌœ «·ﬁ”ÿ  —ﬁ„   " & .Cell(flexcpTextDisplay, Row, .ColIndex("QeqtNum"))
            LogTextE = " Select Instllmen No  " & .Cell(flexcpTextDisplay, Row, .ColIndex("QeqtNum"))
        Else
            LogTextA = " «·€«¡   ÕœÌœ «·ﬁ”ÿ  —ﬁ„   " & .Cell(flexcpTextDisplay, Row, .ColIndex("QeqtNum"))
            LogTextE = " ]ÀDeSelect Instllmen No  " & .Cell(flexcpTextDisplay, Row, .ColIndex("QeqtNum"))
                                                      
        End If
                                                          
         LogTextA = LogTextA & " »ﬁÌ„…   " & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & "   «—ÌŒ «·«” Õﬁ«ﬁ  " & .Cell(flexcpTextDisplay, Row, .ColIndex("Due_Date")) & "   Õ’Ì· ”«»ﬁ  " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " »‰«¡ ⁄·Ï  " & .Cell(flexcpTextDisplay, Row, .ColIndex("TransactionTypeName")) & " »—ﬁ„   " & .Cell(flexcpTextDisplay, Row, .ColIndex("NoteSerial1"))
            LogTextE = LogTextE & " Value " & .Cell(flexcpTextDisplay, Row, .ColIndex("Value")) & " Due Date   " & .Cell(flexcpTextDisplay, Row, .ColIndex("Due_Date")) & " Collection prior " & .Cell(flexcpTextDisplay, Row, .ColIndex("Des")) & " Based On " & .Cell(flexcpTextDisplay, Row, .ColIndex("TransactionTypeName")) & "  No  " & .Cell(flexcpTextDisplay, Row, .ColIndex("NoteSerial1"))
           AddToLogFile CInt(user_id), NoteType, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg
                                                      
        '    End If
                  
        ''''''''''''''''''''''
    End With

    If Me.TxtModFlg.Text = "N" Then
        Me.TxtQastNO.Text = IntCounter
        Me.TxtSum.Text = DblSum - Payment

        If CboResType.ListIndex = 0 Then
            TxtValue.Text = Me.TxtSum.Text
        End If

    Else
        IntCounter = IntCounter + FgDetails.Rows - 1
        DblSum = DblSum + FgDetails.Aggregate(flexSTSum, 1, FgDetails.ColIndex("Value"), FgDetails.Rows - 1, FgDetails.ColIndex("Value"))
        Me.TxtQastNO.Text = IntCounter
        Me.TxtSum.Text = DblSum - Payment
    End If

    CalPre
    Exit Sub
ErrTrap:
End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    If Col <> FG.ColIndex("Released") Then
        Cancel = True
    End If

End Sub

Private Sub ChangeLang()
    'CmdConvert.Caption = "Convert to bill"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
CmdAttach.Caption = "Attachments"

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(19).Caption = "Entry No."
    Cmd(9).Caption = "Print JL "
    Me.TabMain.TabCaption(0) = "Deferred premiums"
    
    Me.TabMain.TabCaption(1) = "Previous installments"

    Me.Caption = "Collection and payment of premiums"
    Me.ELe(6).Caption = Me.Caption

    lbl(5).Caption = "ID"
    lbl(18).Caption = "OPR"

    lbl(0).Caption = "Date"
    lbl(11).Caption = "OPR Type"
    lbl(7).Caption = "Customer"
    lbl(10).Caption = "Cust. Type"
    lbl(9).Caption = "Box"
    lbl(3).Caption = "collection Type"
    lbl(24).Caption = "install. count"
    lbl(25).Caption = "install. Total"
    lbl(6).Caption = "Amount collected"
    lbl(8).Caption = "Premiums in the current OPR"

    With FgDetails
        .TextMatrix(0, .ColIndex("BillID")) = "Bill ID"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("Due_Date")) = "Due Date"
    End With

    With FgReceipted
        .TextMatrix(0, .ColIndex("TransactionTypeName")) = "Transaction Type"
        .TextMatrix(0, .ColIndex("Transaction_Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("QeqtNum")) = "install.#"
        .TextMatrix(0, .ColIndex("Code")) = "Code"
        .TextMatrix(0, .ColIndex("ReceiptType")) = "Receipt Type"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("Due_Date")) = "Due Date"
    End With

    With FgPayed
        .TextMatrix(0, .ColIndex("TransactionTypeName")) = "Transaction Type"
        .TextMatrix(0, .ColIndex("Transaction_Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("QeqtNum")) = "install.#"
        .TextMatrix(0, .ColIndex("Code")) = "Code"
        .TextMatrix(0, .ColIndex("ReceiptType")) = "Receipt Type"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("Due_Date")) = "Due Date"
    End With

    With FG
        .TextMatrix(0, .ColIndex("serial")) = "Index"
        .TextMatrix(0, .ColIndex("Released")) = "Released"
        .TextMatrix(0, .ColIndex("Value")) = "Value"
        .TextMatrix(0, .ColIndex("Due_Date")) = "Due Date"

        .TextMatrix(0, .ColIndex("Des")) = "Des"
        .TextMatrix(0, .ColIndex("QeqtNum")) = "install.#"
        .TextMatrix(0, .ColIndex("TransactionTypeName")) = "Transaction Type"
        .TextMatrix(0, .ColIndex("Transaction_Serial")) = "Serial"

    End With

    ELe(7).Caption = "GL"
    lbl(30).Caption = "GL#"
    lbl(29).Caption = "Interval"
    lbl(32).Caption = "Depit1"
    lbl(16).Caption = "Depit2"
    lbl(31).Caption = "Credit1"
    lbl(17).Caption = "Credit2"
    lbl(1).Caption = " By:"
    lbl(2).Caption = "Curr. Rec."
    lbl(4).Caption = "Rec. Count:"
    lbl(14).Caption = "The outcome of the premiums"
    lbl(15).Caption = "Premiums paid"

    '
 
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.Cmd(8).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    
End Sub

Private Sub Form_Load()
    Dim StrSQL As String
    Dim BGround As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    On Error GoTo ErrTrap

    ScreenNameArabic = " Õ’Ì· Ê”œ«œ «·√ﬁ”«ÿ "
    ScreenNameEnglish = " Pay and Collect Installments"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
    Dim FirstPeriod As Date
    getFirstPeriodDateInthisYear FirstPeriod
    DTPickerAccFrom.value = FirstPeriod
 
    'Resize_Form Me
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    FG.WallPaper = BGround.Picture
    FgDetails.WallPaper = BGround.Picture
    FgReceipted.WallPaper = BGround.Picture
    FgPayed.WallPaper = BGround.Picture
    LoadIcons
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set Cmd(8).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture

    If SystemOptions.UserInterface = EnglishInterface Then

        With Me.CboType
            .Clear
            .AddItem "The collection of premiums"
            .AddItem "Payment of premiums"
        End With

        With Me.CboDealerType
            .Clear
            .AddItem "Customer"
            .AddItem "Vendor"
        End With

        With CboResType
            .Clear
            .AddItem "Normal collection"
            .AddItem "Collection at a discount"
            .AddItem "Part of the premium"
            .AddItem "Adv. Payment"
            .AddItem "Paid from the account  "
             
        End With

        With CboPrecenType
            .AddItem "Percentage  ", 0
            .ItemData(0) = 1
            .AddItem "  Fixed Value", 1
            .ItemData(1) = 2
            .AddItem "No Discount", 2
            .ItemData(2) = 3
            .ListIndex = 2
        End With

    Else

        With CboPrecenType
            .AddItem "‰”»… „∆ÊÌ…", 0
            .ItemData(0) = 1
            .AddItem "ﬁÌ„… À«» …", 1
            .ItemData(1) = 2
            .AddItem "·«ÌÊÃœ", 2
            .ItemData(2) = 3
            .ListIndex = 2
        End With

        With Me.CboType
            .Clear
            .AddItem " Õ’Ì·  "
            .AddItem "”œ«œ  "
        End With

        With Me.CboDealerType
            .Clear
            .AddItem "⁄„Ì·"
            .AddItem "„Ê—œ"
        End With

        With CboResType
            .Clear
            .AddItem " Õ’Ì· ⁄«œÌ"
            .AddItem " Õ’Ì· »Œ’„"
            .AddItem "œ›⁄… „‰   «·ﬁ”ÿ"
            .AddItem "œ›⁄…     „ﬁœ„Â "
            .AddItem "œ›⁄… „‰   «·Õ”«» "
        End With

    End If

    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, False
    Dcombos.GetBranches Me.Dcbranch

    Set cDcboSearch(0) = New clsDCboSearch
    Set cDcboSearch(0).Client = Me.DBCboClientName
    Set cDcboSearch(1) = New clsDCboSearch
    Set cDcboSearch(1).Client = Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName

    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboDebitSide1
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide1
    Dcombos.GetBranches Me.Dcbranch
   Dcombos.GetAccountingCodes Me.DCAccounts, True
 
    If SystemOptions.usertype <> UserAdminAll Then
        Me.Dcbranch.Enabled = True
    End If

    With Me.CboPayMentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "ÕÊ«·Â »‰ﬂÌÂ"
        .AddItem "‘Ìﬂ  „Õ’·"
        .AddItem "Õ”«»"
    End With
 
    Dcombos.GetChequeBox Me.DCChequeBox

    Dcombos.GetBanks Me.DcboBankName

    AddTip
    Set rs = New ADODB.Recordset
    rs.Open "[ReceiptQest]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Me.TxtModFlg.Text = "R"
    TabMain.CurrTab = 0
    SetDtpickerDate Me.DtbBill
   Resize_Form Me, TransactionSize
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Function saveChequeBoxContents(NoteID As Double)

    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords
 
    If val(DCChequeBox.BoundText) = 0 Then Exit Function

   ' rs.Open "TblChecqueBoxContent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     * from dbo.TblChecqueBoxContent Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    rs.AddNew
    rs("noteid").value = NoteID
    rs("ChequeBoxID").value = val(DCChequeBox.BoundText)
            
    ' rs("RecordDate").value = XPDtbTrans.value
    rs("DueDate").value = DtpChequeDueDate.value
    rs("BankName").value = TXTBankName.Text
    rs("ChequeNo").value = TxtChequeNumber.Text
    rs("ChequeValue").value = val(TxtValue.Text)
    
    rs("Remarks").value = DcboCreditSide.Text
    rs("Deposited").value = 0
    rs("Collected").value = 0
    rs("CreditAccount").value = (DcboCreditSide.BoundText)
    rs.update
  
    rs.Close
End Function

 

Private Sub Txt_Change(Index As Integer)
    CalPre
End Sub

Private Sub TxtCustCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtCustCode.Text
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            '     Me.Caption = " Õ’Ì· Ê”œ«œ «·√ﬁ”«ÿ"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            Me.txtid.locked = True
            Me.DBCboClientName.locked = True
            TxtValue.locked = True
            Me.DtbBill.Enabled = False
            Me.CboType.locked = True
            Me.CboDealerType.locked = True
            CboResType.locked = True
            FG.Editable = flexEDNone
            Me.Txt_akchen.locked = True

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

            Me.DcboBox.locked = True

        Case "N"
            '     Me.Caption = " Õ’Ì· Ê”œ«œ «·√ﬁ”«ÿ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
        
            Me.txtid.locked = False
            Me.DBCboClientName.locked = False
            Me.DtbBill.Enabled = True
            Me.DtbBill.value = Date
            Me.CboType.locked = False
            Me.CboDealerType.locked = False
            TxtValue.locked = False
            CboResType.locked = False
            FG.Editable = flexEDKbdMouse
            CboResType.ListIndex = 0
            Me.DcboBox.locked = False
            Me.Txt_akchen.locked = False

        Case "E"
            '     Me.Caption = " Õ’Ì· Ê”œ«œ «·√ﬁ”«ÿ(  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            Me.txtid.locked = False
            Me.DBCboClientName.locked = False
            Me.DtbBill.Enabled = True
            Me.CboType.locked = False
            Me.CboDealerType.locked = False
            TxtValue.locked = False
            CboResType.locked = False
            FG.Editable = flexEDKbdMouse
            Me.DcboBox.locked = False
            Me.Txt_akchen.locked = False

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtValue_Change()
    Me.Lb_note_value_by_characters.Caption = WriteNo(Format(Me.TxtValue.Text, "0.00"), 0, True, ".")
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtValue.Text, 0)
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

Function CalcNoOfInstallments(SumValue As Double) As Boolean
    On Error Resume Next
    Dim RowNum As Integer
    Dim TempRow As Integer
    Dim QestValue As Double
    Dim Reminder As Double
    Dim RowValue As Double
    Dim totalRowsValue As Double
    Dim CurrentMonth  As String
    CalcNoOfInstallments = True
    Reminder = SumValue
    totalRowsValue = 0
    CurrentMonth = ""

    For TempRow = FG.FixedRows To FG.Rows - 1

        If FG.Rowdata(TempRow) <> "" Then
            FG.Cell(flexcpChecked, TempRow, FG.ColIndex("Released")) = flexUnchecked
            FG.TextMatrix(TempRow, FG.ColIndex("CurrentValue")) = ""
        End If

    Next TempRow

    For TempRow = FG.FixedRows To FG.Rows - 1
        RowValue = (val(FG.TextMatrix(TempRow, FG.ColIndex("Value"))) - val(FG.TextMatrix(TempRow, FG.ColIndex("Des"))))
        
        If Reminder >= RowValue Then
            FG.TextMatrix(TempRow, FG.ColIndex("CurrentValue")) = RowValue
            Reminder = Reminder - RowValue
            FG.Cell(flexcpChecked, TempRow, FG.ColIndex("Released")) = flexChecked
            CurrentMonth = CurrentMonth & " Ê" & MonthName(Month(FG.TextMatrix(TempRow, FG.ColIndex("Due_Date"))))
                          
        ElseIf Reminder < RowValue And Reminder > 0 Then
            FG.TextMatrix(TempRow, FG.ColIndex("CurrentValue")) = Reminder
            FG.Cell(flexcpChecked, TempRow, FG.ColIndex("Released")) = flexChecked
            CurrentMonth = CurrentMonth & " Ê Ã“¡ „‰ " & MonthName(Month(FG.TextMatrix(TempRow, FG.ColIndex("Due_Date"))))
            Exit For
        End If
          
        totalRowsValue = totalRowsValue + RowValue

        If totalRowsValue >= SumValue Then
            Exit For
        End If

    Next TempRow

    CurrentMonth = mId(CurrentMonth, 3, Len(CurrentMonth) - 1)
    Me.TXTMessageDES.Text = CurrentMonth
    Dim Msg As String
 
    'Dim RowNum As Integer
    Dim IntCounter As Integer
    Dim DblSum As Double
    Dim NoOfQest As Integer
    DblSum = 0
    NoOfQest = 0

    With FG

        For RowNum = .FixedRows To .Rows - 1

            If .Cell(flexcpChecked, RowNum, .ColIndex("Released")) = flexChecked Then
                NoOfQest = NoOfQest + 1
                DblSum = DblSum + val(.TextMatrix(RowNum, .ColIndex("CurrentValue")))
 
            End If

        Next RowNum

    End With
 
    TxtQastNO.Text = NoOfQest
    TxtSum.Text = DblSum
 
    If DblSum < SumValue Then
        Msg = "«·ﬁÌ„… «·„œ›Ê⁄Â «ﬂ»— „‰ «Ã„«·Ì «·«ﬁ”«ÿ ÌÃ» „—«Ã⁄Â ﬁÌ„… «·œ›⁄Â «·„ﬁœ„… "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
     
        CalcNoOfInstallments = False
   
    End If

End Function

Private Sub DCAccounts_Click(Area As Integer)
    DCAccounts_Change
End Sub
Private Sub DCAccounts_Change()

    If DCAccounts.BoundText = "" Then Exit Sub
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        DcboDebitSide.BoundText = DCAccounts.BoundText
    End If

End Sub



Private Sub DCAccounts_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 23072014
    End If

End Sub
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim BeginTrans As Boolean
    Dim TempRow As Integer
    Dim SngTemp As Single
    Dim SngDisDiff As Single
    Dim SngOnePartDis As Single
    Dim SngOnePartValue As Single
    Dim LngDevID As Long
    Dim RsDev As ADODB.Recordset
    Dim IntDevLineNO As Integer
    Dim IntResult As String

   On Error GoTo ErrTrap

    If Me.CboType.ListIndex = -1 Then
        Msg = "ÌÃ» ‰Ê⁄ «·⁄„Ì·… Â· (  Õ’Ì· √ﬁ”«ÿ «„  ”œÌœ √ﬁ”«ÿ)..!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboType.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If DBCboClientName.BoundText = "" Then
        Msg = "ÌÃ»  ÕœÌœ «”„ «·⁄„Ì·"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DBCboClientName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If CboResType.ListIndex <> 3 And CboResType.ListIndex <> 4 Then
        If val(Trim(Me.TxtQastNO.Text)) <= 0 Then
            Msg = "ÌÃ» ≈Œ Ì«— ﬁ”ÿ Ê«Õœ ⁄·Ï «·√ﬁ· " & CHR(13)
            Msg = Msg & "·Ì „  Õ’Ì·Â...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If

    If CboResType.ListIndex < 0 Then
        Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «· Õ’Ì·"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboResType.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    'If Me.DcboBox.BoundText = "" Then
    '    Msg = "ÌÃ»  ÕœÌœ √”„ «·Œ“‰…...!!!"
    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    DcboBox.SetFocus
    '    SendKeys "{F4}"
    '    Exit Sub
    'End If

    If Me.CboPayMentType.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄...!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboPayMentType.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If Me.CboPayMentType.ListIndex = 0 Then
        If Me.DcboBox.BoundText = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboBox.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

    ElseIf Me.CboPayMentType.ListIndex = 1 Then
      
        '  If DateDiff("d", Me.DtpChequeDueDate.value, Date) > 0 Then
        '      Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
        '      MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '      DtpChequeDueDate.SetFocus
        '      SendKeys "{F4}"
        '      Exit Sub
        '  End If
        If SystemOptions.ChequeBox = True Then
         
            If DCChequeBox.BoundText = "" Then
                           
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Õœœ Õ«›Ÿ… «·‘Ìﬂ«  ...!!"
                Else
                    Msg = "Select Cheque Box ...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DCChequeBox.SetFocus
                SendKeys "{F4}"
                Exit Sub
                   
            End If
    
            If TXTBankName.Text = "" Then
                           
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "«ﬂ » «”„ »‰ﬂ «·‘Ìﬂ    « ...!!"
                Else
                    Msg = " Enter Bank Name For Cheque  ...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TXTBankName.SetFocus
                SendKeys "{F4}"
                Exit Sub
                    
            End If
        
            If Trim$(Me.TxtChequeNumber.Text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If

        Else
       
            If Me.DcboBankName.BoundText = "" Then
                Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcboBankName.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.Text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                TxtChequeNumber.SetFocus
                Exit Sub
            End If
        End If
    
    ElseIf Me.CboPayMentType.ListIndex = 2 Then

        If Me.DcboBankName.BoundText = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboBankName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If Trim$(Me.TxtChequeNumber.Text) = "" Then
            Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·ÕÊ«·Â...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtChequeNumber.SetFocus
            Exit Sub
        End If
     
    ElseIf Me.CboPayMentType.ListIndex = 3 Then

        If Me.DcboBankName.BoundText = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboBankName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If Trim$(Me.TxtChequeNumber.Text) = "" Then
            Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtChequeNumber.SetFocus
            Exit Sub
        End If
     
     
   ElseIf Me.CboPayMentType.ListIndex = 4 Then
            If Trim(Me.DCAccounts.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— ·Õ”«»..!!"
                Else
                    Msg = "Select Account..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DCAccounts.SetFocus
                SendKeys "{F4}"
                Exit Sub
            End If
        
        End If
  
     
    
    '       If Me.DCChequeBox.BoundText <> "" Then
    '   If ChequeBoxOperations(Val(Me.XPTxtID)) = False Then
    '       Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
    '       Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
    '       MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '       Exit Sub
    '   End If
    'End If

    If Me.CboType.ListIndex = 1 Then
        '    If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.TxtValue.text), Me.DtbBill.value, False) = False Then
        '        Msg = "⁄›Ê« —’Ìœ «·Œ“‰… «·„Õœœ… ·« Ì”„Õ »≈ „«„ «·⁄„·Ì…"
        '        Msg = Msg & Chr(13) & "»—Ã«¡ „—«Ã⁄… —’Ìœ «·Œ“‰… «·„Õœœ…"
        '        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        DcboBox.SetFocus
        '        SendKeys "{F4}"
        '        Exit Sub
        '    End If
    End If

    If val(Me.TxtValue.Text) = 0 And CboResType.ListIndex <> 1 Then
        Msg = "ÌÃ» ﬂ «»… ﬁÌ„… «·„»·€ «·„Õ’·"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtValue.SetFocus
        Exit Sub
    End If

    'If Val(Me.Txt_akchen.text) = 0 Then
    '    Msg = "Â·  —Ìœ ﬂ «»… —ﬁ„ «·Õ—ﬂ… "
    '   MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    '  IntResult = MsgBox(Msg, vbMsgBoxRight + vbYesNo + vbMsgBoxRtlReading + vbQuestion, App.Title)

    'Select Case IntResult
    '    Case vbYes
    '
    '    Txt_akchen.SetFocus
    '    Exit Sub
    '    Case vbNo
    '
    'End Select
    'End If

    'With CboResType
    '    .AddItem " Õ’Ì· ⁄«œÌ"
    '    .AddItem " Õ’Ì· »Œ’„"
    '    .AddItem "œ›⁄… „‰ Õ”«» «·ﬁ”ÿ"
    'End With
    For TempRow = FG.FixedRows To FG.Rows - 1

        If FG.Rowdata(TempRow) <> "" Then
            If FG.Cell(flexcpChecked, TempRow, FG.ColIndex("Released")) = flexChecked Then
                Exit For
            End If
        End If

    Next TempRow

    SngTemp = 0

    If TempRow = FG.Rows Then
        SngTemp = 0
    Else

        SngTemp = Before_Release(val(FG.Rowdata(TempRow)))
    End If

    If CboResType.ListIndex = 0 Then ' Õ’Ì· ⁄«œÌ
        If SngTemp <> 0 Then
            Msg = "Â–« «·ﬁ”ÿ ﬁœ  „  Õ’Ì· Ã“¡ „‰Â ”«»ﬁ«" & CHR(13)
            Msg = Msg & "Ê·–« ·«Ì„ﬂ‰  Õ’Ì·Â ﬂ Õ’Ì· ⁄«œÏ!!!" & CHR(13)
            Msg = Msg & "·–« Ì—ÃÏ  €Ì— ‰Ê⁄ «· Õ’Ì· ≈·Ï --( œ›⁄… „‰ Õ”«»  )" & CHR(13)
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboResType.SetFocus
            CboResType.ListIndex = 4
            Exit Sub
        End If

        If val(TxtValue.Text) <> val(Me.TxtSum.Text) Then
            Msg = "›Ï Õ«·… «· Õ’Ì· «·⁄«œÏ ·«»œ Ê√‰ Ì ”«ÊÏ" & CHR(13)
            Msg = Msg & " ﬁÌ„… «·„»·€ «·„Õ’· „⁄ ﬁÌ„… ≈Ã„«·Ï «·√ﬁ”«ÿ " & CHR(13)
            Msg = Msg & "»—Ã«¡  ⁄œÌ· ﬁÌ„… «·„»·€ «·„Õ’·" & CHR(13)
            Msg = Msg & "«Ê  ⁄œÌ· ‰Ê⁄ «· Õ’Ì·"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '        TxtValue.SetFocus
            Exit Sub
        End If

    ElseIf CboResType.ListIndex = 1 Then

        If val(TxtValue.Text) >= val(Me.TxtSum.Text) Then
            Msg = "⁄‰œ„« ÌﬂÊ‰ ‰Ê⁄ «· Õ’Ì· .. Õ’Ì· »Œ’„" & CHR(13)
            Msg = Msg & "·«»œ «‰ ÌﬂÊ‰ «·„»·€ «·„Õ’· √ﬁ· „‰ ≈Ã„«·Ï «·√ﬁ”«ÿ"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    
        If CboPrecenType.ListIndex = 0 Or CboPrecenType.ListIndex = 1 Then
            If Txt(3).Text = "" Then
                Msg = "›Ì Õ«·… ÊÃÊœ ›«∆œ… ÌÃ»  ÕœÌœ ﬁÌ„… √Ê ‰”»… Â–Â «·›«∆œ…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

                If Txt(3).Enabled = True Then
                    Txt(3).SetFocus
                End If

                Exit Sub
            End If

            If Not IsNumeric(Txt(3).Text) Then
                Msg = " ﬁÌ„… √Ê ‰”»… Â–Â «·›«∆œ… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

                If Txt(3).Enabled = True Then
                    Txt(3).SetFocus
                End If

                Exit Sub
            End If
        End If

        SngDisDiff = val(Me.TxtSum.Text) - val(TxtValue.Text)
        SngOnePartDis = SngDisDiff / val(Me.TxtQastNO.Text)
    
    ElseIf CboResType.ListIndex = 2 Then 'œ›⁄… „‰ Õ”«» «·ﬁ”ÿ

        If val(TxtQastNO.Text) > 1 Then
            Msg = "⁄‰œ  Õ’Ì· œ›⁄…  Õ  Õ”«» «·ﬁ”ÿ " & CHR(13)
            Msg = Msg & "·«»œ Ê«‰ ÌﬂÊ‰ ﬁ”ÿ Ê«Õœ ›ﬁÿ " & CHR(13)
            Msg = Msg & "»—Ã«¡ „—«Ã⁄… «·√ﬁ”«ÿ «·„Õ’·… «Ê  €Ì— " & CHR(13)
            Msg = Msg & "‰Ê⁄ «· Õ’Ì·...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        ElseIf val(TxtValue.Text) >= val(FG.TextMatrix(TempRow, FG.ColIndex("Value"))) Then
            Msg = "⁄‰œ  Õ’Ì· œ›⁄…  Õ  Õ”«» «·ﬁ”ÿ " & CHR(13)
            Msg = Msg & "·«»œ Ê«‰ ÌﬂÊ‰ ﬁÌ„… «·„»·€ «·„œ›Ê⁄ " & CHR(13)
            Msg = Msg & "«ﬁ· „‰ ﬁÌ„… «·ﬁ”ÿ " & CHR(13)
            Msg = Msg & "»—Ã«¡ „—«Ã⁄… «·ﬁÌ„… «·„œŒ·… "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        Else
            'Code here to calculate the before release
            SngTemp = 0
            SngTemp = Before_Release(val(FG.Rowdata(TempRow)))

            If val(Me.TxtValue) + SngTemp > val(FG.TextMatrix(TempRow, FG.ColIndex("Value"))) Then
                Msg = "Â–« «·„»·€ «·„Õ’· €Ì— „ﬁ»Ê· " & CHR(13)
                Msg = Msg & "ÕÌÀ √‰ Â‰«ﬂ œ›⁄«  „Õ’·… „”»ﬁ« " & CHR(13)
                Msg = Msg & "„‰ Â–« «·ﬁ”ÿ " & CHR(13)
                Msg = Msg & "≈Ã„«·Ï «·œ›⁄«  «·„Õ’·… „”»ﬁ« = " & SngTemp
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If
        End If
    
    ElseIf CboResType.ListIndex = 3 Or CboResType.ListIndex = 4 Then  'œ›⁄…     „ﬁœ„Â
    
        If val(TxtValue.Text) = 0 Then
            Msg = "›Ï Õ«·…  «·œ›⁄Â «·„ﬁœ„Â «Ê œ›⁄Â „‰ «·Õ”«»    ·«»œ  „‰ ﬂ «»… " & CHR(13)
            Msg = Msg & " ﬁÌ„… «·„»·€ «·„Õ’· " & CHR(13)
        
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '        TxtValue.SetFocus
            Exit Sub
        End If
    
        If CalcNoOfInstallments(val(Me.TxtValue)) = False Then
            Exit Sub
        Else
     
        End If
    
    End If
                       
    Dim notes_result As String
    Dim Vchr_result As String

    If TxtNoteSerial1.Text = "" Then
        Vchr_result = Voucher_coding(val(my_branch), DtbBill.value, Sanad_No, notesType)

        If Vchr_result = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ    Õ’Ì· «ﬁ”«ÿ ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        Else
                       
            If Vchr_result = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            Else
                '         txtNoteSerial1.text = Voucher_coding(val(my_branch), DtbBill.value, Sanad_No, notesType)
            End If
        End If
    End If
                       
    If TxtNoteSerial.Text = "" Then
        notes_result = Notes_coding(val(my_branch), DtbBill.value)

        If notes_result = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else
                       
            If notes_result = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                'TxtNoteSerial.text = Notes_coding(val(my_branch), DtbBill.value)
            End If
        End If
    End If

    Cn.BeginTrans
    BeginTrans = True

    Select Case Me.TxtModFlg.Text

        Case "N"
            rs.AddNew
            txtid.Text = CStr(new_id("ReceiptQest", "ReceiptID", "", True))
            XPTxtID.Text = CStr(new_id("Notes", "NoteID", "", True))

            rs("ReceiptID").value = val(txtid.Text)
            Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)

        Case "E"
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where ReceiptID=" & val(txtid.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
 
            StrSQL = "Delete From notes Where NoteID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adCmdText
        
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where NOTES_ID=" & val(Me.XPTxtID.Text)
            Cn.Execute StrSQL, , adCmdText
        
    End Select

    rs("Cust_ID").value = IIf(DBCboClientName.BoundText = "", Null, (DBCboClientName.BoundText))
    rs("ReceiptDate").value = DtbBill.value
    rs("ReceiptDateH").value = ToHijriDate(DtbBill.value)
    If CboResType.ListIndex > -1 Then
        rs("ReceiptType").value = CboResType.ListIndex
    End If

    If CboResType.ListIndex = 1 Then
        rs("DiscountType").value = val(CboPrecenType.ListIndex)
        rs("DiscounVal").value = val(Txt(3).Text)
    Else
        rs("DiscountType").value = Null
        rs("DiscounVal").value = Null
    End If

    rs("Remark").value = IIf(txtRemark.Text = "", Null, txtRemark.Text)
    rs("MessageDES").value = IIf(TXTMessageDES.Text = "", Null, TXTMessageDES.Text)

    rs("PartCount").value = IIf(TxtQastNO.Text = "", 0, TxtQastNO.Text)
    rs("Total").value = IIf(TxtSum.Text = "", 0, TxtSum.Text)
    rs("PaymentMoney").value = IIf(TxtValue.Text = "", "", TxtValue.Text)
    rs("User_ID").value = IIf(DCboUserName.BoundText = "", Null, (DCboUserName.BoundText))
    'rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, (DcboBox.BoundText))

    rs("BankName").value = IIf(TXTBankName.Text = "", "", Trim(TXTBankName.Text))

    'ÿ—Ìﬁ… «·œ›⁄ «·‰ﬁœÏ «Ê «·‘Ìﬂ
    If Me.CboPayMentType.ListIndex = 0 Then
        rs("NoteCashingType").value = 0
        rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, DcboBox.BoundText)
        rs("BankID").value = Null
        rs("ChqueNum").value = Null
        rs("DueDate").value = Null
        
    ElseIf Me.CboPayMentType.ListIndex = 1 Then
        rs("NoteCashingType").value = 1
        rs("BoxID").value = Null

        If SystemOptions.ChequeBox = False Then
        
            rs("BankID").value = val(Me.DcboBankName.BoundText)
        End If
        
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
        rs("DueDate").value = Me.DtpChequeDueDate.value

        If SystemOptions.ChequeBox = True Then
            rs("ChequeBoxID").value = IIf(DCChequeBox.BoundText = "", Null, DCChequeBox.BoundText)
        Else
            rs("ChequeBoxID").value = Null
                
        End If
                
    ElseIf Me.CboPayMentType.ListIndex = 2 Then
        rs("NoteCashingType").value = 2
        rs("BoxID").value = Null
        rs("BankID").value = val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("ChequeBoxID").value = Null
                
    ElseIf Me.CboPayMentType.ListIndex = 3 Then
        rs("NoteCashingType").value = 3
        rs("BoxID").value = Null
        rs("BankID").value = val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.Text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("ChequeBoxID").value = Null
               
          ElseIf Me.CboPayMentType.ListIndex = 4 Then
        rs("BoxID").value = Null
        rs("BankID").value = Null
        rs("ChqueNum").value = Null
        rs("DueDate").value = Null
        rs("NoteCashingType").value = 4
        rs("AccountCode").value = (Me.DCAccounts.BoundText)
        
 
        
    End If
    
     

    '--------

    rs("NumAkch").value = IIf(Txt_akchen.Text = "", 0, Txt_akchen.Text)

    If Me.CboType.ListIndex = 0 Then
        ' Õ’Ì· ﬁ”ÿ „” Õﬁ ··‘—ﬂ…
        rs("OperationType").value = 0
    ElseIf CboType.ListIndex = 1 Then
        '”œ«œ ﬁ”ÿ ⁄·Ï «·‘—ﬂ…
        rs("OperationType").value = 1
    End If

    rs("NoteID").value = IIf(Me.XPTxtID.Text = "", Null, (Me.XPTxtID.Text))

    If TxtNoteSerial.Text = "" Then
        TxtNoteSerial.Text = Notes_coding(val(my_branch), DtbBill.value)
    End If
                                   
    If TxtNoteSerial1.Text = "" Then
        TxtNoteSerial1.Text = Voucher_coding(val(my_branch), DtbBill.value, Sanad_No, notesType)
    End If
              
    Set RsNotes = New ADODB.Recordset
    'RsNotes.Open "[notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    RsNotes.AddNew
    RsNotes("branch_no").value = val(Me.Dcbranch.BoundText)
    RsNotes("NoteID").value = val(XPTxtID.Text)
    RsNotes("NoteSerial").value = IIf(Me.TxtNoteSerial.Text = "", Null, Me.TxtNoteSerial.Text)
    RsNotes("NoteSerial1").value = IIf(Me.TxtNoteSerial1.Text = "", Null, (Me.TxtNoteSerial1.Text))
              
    rs("NoteSerial").value = IIf(Me.TxtNoteSerial.Text = "", Null, (Me.TxtNoteSerial.Text))
    rs("NoteSerial1").value = IIf(Me.TxtNoteSerial1.Text = "", Null, (Me.TxtNoteSerial1.Text))
    rs("OldNoteSerial1").value = IIf(Me.oldtxtNoteSerial1.Text = "", Null, (Me.oldtxtNoteSerial1.Text))
    'rs("branch_no").value = Val(Me.Dcbranch.BoundText)
    rs("branch_no").value = val(Me.Dcbranch.BoundText)
    rs.update
   ' RsTemp.Open "InstallmentDet_Junc_Receipt", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     StrSQL = "SELECT     * from dbo.InstallmentDet_Junc_Receipt Where (1 = -1)"
   RsTemp.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
    For RowNum = 1 To FG.Rows - 1

        If FG.Rowdata(RowNum) <> "" Then
            If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("Released")) = flexChecked Then
                RsTemp.AddNew
                RsTemp("JuncID").value = CStr(new_id("InstallmentDet_Junc_Receipt", "JuncID", "", True))
                RsTemp("QestID").value = FG.Rowdata(RowNum)
                RsTemp("ReceiptID").value = val(txtid.Text)

                If CboResType.ListIndex = 0 Then
                    RsTemp("Status").value = 0
                    RsTemp("Value").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Value")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Value")) - val(FG.TextMatrix(RowNum, FG.ColIndex("Des"))))
                ElseIf CboResType.ListIndex = 1 Then ' Õ’Ì· »Œ’„
                    RsTemp("Status").value = 0

                    'SngOnePartDis
                    If val(FG.TextMatrix(RowNum, FG.ColIndex("Value"))) = 0 Then
                        SngOnePartValue = 0
                    Else
                        SngOnePartValue = val(FG.TextMatrix(RowNum, FG.ColIndex("Value"))) - (val(FG.TextMatrix(RowNum, FG.ColIndex("Des"))) + SngOnePartDis)
                    End If

                    RsTemp("Value").value = SngOnePartValue
                ElseIf CboResType.ListIndex = 2 Then
                    RsTemp("Status").value = 1
                    RsTemp("Value").value = IIf(TxtValue.Text = "", 0, val(TxtValue.Text))
                ElseIf CboResType.ListIndex = 3 Or CboResType.ListIndex = 4 Then

                    If val(FG.TextMatrix(RowNum, FG.ColIndex("CurrentValue"))) = val(FG.TextMatrix(RowNum, FG.ColIndex("Value"))) - val(FG.TextMatrix(RowNum, FG.ColIndex("Des"))) Then
                        RsTemp("Status").value = 0
                    Else
                        RsTemp("Status").value = 1
                     
                    End If

                    RsTemp("Value").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("CurrentValue")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("CurrentValue")))
                    '    RsTemp("CurrentValue").value = IIf(Fg.TextMatrix(RowNum, Fg.ColIndex("CurrentValue")) = "", "", Fg.TextMatrix(RowNum, Fg.ColIndex("CurrentValue")))
                End If

                RsTemp.update
            End If
        End If

    Next RowNum

    '==========================================================================
    ' ”ÃÌ· notes
        
    saveChequeBoxContents (XPTxtID.Text)
    
    RsNotes("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
    RsNotes("Note_Value").value = IIf(TxtValue.Text = "", Null, val(TxtValue.Text))
     
    If CboResType.ListIndex = 1 Then
        Me.Lb_note_value_by_characters.Caption = WriteNo(Format(Me.TxtSum.Text, "0.00"), 0, True, ".")
             
    Else
        Me.Lb_note_value_by_characters.Caption = WriteNo(Format(Me.TxtValue.Text, "0.00"), 0, True, ".")
    End If
    
    RsNotes("note_value_by_characters").value = Trim$(Me.Lb_note_value_by_characters.Caption)
    
    RsNotes("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ

    If CboType.ListIndex = 0 Then
        RsNotes("numbering_type1").value = sand_numbering_type(25)  '”‰œ   Õ’Ì·
    Else
        RsNotes("numbering_type1").value = sand_numbering_type(26) '”‰œ  ”œ·œ
    End If
    
    RsNotes("sanad_year").value = year(DtbBill.value)
    RsNotes("sanad_month").value = Month(DtbBill.value)
     
    '  CboResType.text & " ·  " & Me.DBCboClientName.text & chr(13) & TxtRemark.Text
    RsNotes("Remark").value = IIf(CboType.Text = "", "", Trim(CboType.Text)) & CboResType.Text & " · " & DBCboClientName.Text & CHR(13) & txtRemark.Text
    RsNotes("BankID").value = Null
    RsNotes("CusID").value = Null

    If CboType.ListIndex = 0 Then
        RsNotes("NoteType").value = 18
    ElseIf CboType.ListIndex = 1 Then
        RsNotes("NoteType").value = 19
    End If

    RsNotes("NoteDate").value = DtbBill.value
    RsNotes("UserID").value = user_id
    RsNotes("BoxID").value = IIf(Me.DcboBox.BoundText = "", Null, Me.DcboBox.BoundText)
    RsNotes("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    RsNotes("sanad_year").value = year(DtbBill.value)
    RsNotes("sanad_month").value = Month(DtbBill.value)
    
    RsNotes.update
    
    ' ”ÃÌ· ﬁÌÊœ
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
        IntDevLineNO = 0
        Set RsDev = New ADODB.Recordset
      '  RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
        '«·ÿ—› «·„œÌ‰
        If val(Me.TxtValue.Text) > 0 Then
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            IntDevLineNO = IntDevLineNO + 1
            RsDev("DEV_ID_Line_No").value = IntDevLineNO
            RsDev("Account_Code").value = Me.DcboDebitSide.BoundText
            RsDev("Value").value = val(Me.TxtValue.Text)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = CboResType.Text & " ·  " & Me.DBCboClientName.Text & CHR(13) & txtRemark.Text
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("ReceiptID").value = val(txtid.Text)
            RsDev("RecordDate").value = Me.DtbBill.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("Double_Entry_Vouchers_Description").value = " ”‰œ  Õ’Ì· «ﬁ”«ÿ »—ﬁ„ :" & TxtNoteSerial1.Text & CHR(13) & CboResType.Text & " ·  " & Me.DBCboClientName.Text & CHR(13) & txtRemark.Text
            RsDev.update
        End If

        If Me.DcboDebitSide1.BoundText <> "" And (val(Me.TxtSum) - val(Me.TxtValue.Text)) > 0 Then
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            IntDevLineNO = IntDevLineNO + 1
            RsDev("DEV_ID_Line_No").value = IntDevLineNO
            RsDev("Account_Code").value = Me.DcboDebitSide1.BoundText
            RsDev("Value").value = val(Me.TxtSum) - val(Me.TxtValue.Text)
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = " ”‰œ  Õ’Ì· «ﬁ”«ÿ »—ﬁ„ :" & TxtNoteSerial1.Text & CHR(13) & CboResType.Text & " ·  " & Me.DBCboClientName.Text & CHR(13) & txtRemark.Text
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("ReceiptID").value = val(txtid.Text)
            RsDev("RecordDate").value = Me.DtbBill.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
        End If

        '«·ÿ—› «·œ«∆‰
        RsDev.AddNew
        RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
        RsDev("Double_Entry_Vouchers_ID").value = LngDevID
        IntDevLineNO = IntDevLineNO + 1
        RsDev("DEV_ID_Line_No").value = IntDevLineNO
        RsDev("Account_Code").value = Me.DcboCreditSide.BoundText

        If CboResType.ListIndex = 1 Then
            RsDev("Value").value = val(Me.TxtSum.Text)
        Else
            RsDev("Value").value = val(Me.TxtValue.Text)
        End If
        
        RsDev("Credit_Or_Debit").value = 1
        RsDev("Double_Entry_Vouchers_Description").value = " ”‰œ  Õ’Ì· «ﬁ”«ÿ »—ﬁ„ :" & TxtNoteSerial1.Text & CHR(13) & CboResType.Text & " ·  " & Me.DBCboClientName.Text & CHR(13) & txtRemark.Text
        RsDev("Notes_ID").value = val(XPTxtID.Text)
        RsDev("ReceiptID").value = val(txtid.Text)
        RsDev("RecordDate").value = Me.DtbBill.value
        RsDev("UserID").value = Me.DCboUserName.BoundText
        RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
        RsDev.update

        If Me.DcboCreditSide1.BoundText <> "" Then
            RsDev.AddNew
            RsDev("branch_id").value = val(Me.Dcbranch.BoundText)
            RsDev("Double_Entry_Vouchers_ID").value = LngDevID
            IntDevLineNO = IntDevLineNO + 1
            RsDev("DEV_ID_Line_No").value = IntDevLineNO
            RsDev("Account_Code").value = Me.DcboCreditSide1.BoundText
            RsDev("Value").value = val(Me.TxtSum) - val(Me.TxtValue.Text)
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = " ”‰œ  Õ’Ì· «ﬁ”«ÿ »—ﬁ„ :" & TxtNoteSerial1.Text & CHR(13) & CboResType.Text & " ·  " & Me.DBCboClientName.Text & CHR(13) & txtRemark.Text
            RsDev("Notes_ID").value = val(XPTxtID.Text)
            RsDev("ReceiptID").value = val(txtid.Text)
            RsDev("RecordDate").value = Me.DtbBill.value
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev.update
        End If

        LblDevID.Caption = LngDevID
        lbl(33).Caption = SystemOptions.SysCurrentAccountIntervalID
    End If

    '==========================================================================
    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    CuurentLogdata

    Select Case Me.TxtModFlg.Text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End Select

    Me.TxtModFlg.Text = "R"
    SendMessage (1)
    Retrive txtid.Text
    DBCboClientName_Change
    lbl(30).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    TxtModFlg.Text = "R"
    'End If
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

Private Function Before_Release(LngQestID As Long) As Single
    Dim StrSQL As String
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrTrap
    StrSQL = "SELECT InstallMentDetails.QestID, Sum(InstallmentDet_Junc_Receipt.Value) " & " AS SumValue "
    StrSQL = StrSQL + " FROM InstallMent INNER JOIN (InstallMentDetails INNER JOIN " & "InstallmentDet_Junc_Receipt ON InstallMentDetails.QestID = " & "InstallmentDet_Junc_Receipt.QestID) ON InstallMent.PartID = InstallMentDetails.PartID "
    StrSQL = StrSQL + " Where InstallMentDetails.QestID=" & LngQestID & ""
    StrSQL = StrSQL + " GROUP BY InstallMentDetails.QestID "
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.BOF Then
        Before_Release = 0
    Else
        Before_Release = IIf(IsNull(rs("SumValue").value), 0, rs("SumValue").value)
    End If

    rs.Close
    Set rs = Nothing
    Exit Function
ErrTrap:
    Before_Release = 0
End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            ClearMe
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "ReceiptID='" & val(txtid.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_TransAction()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If txtid.Text <> "" Then
        '    If Me.CboType.ListIndex = 0 Then
        '        If CheckBoxAccount(Me.DcboBox.BoundText, Val(Me.TxtValue.text), Date, False) = False Then
        '            Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
        '            Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
        '            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '            Exit Sub
        '        End If
        '    End If

        '     If Me.DCChequeBox.BoundText <> "" Then
        '     If ChequeBoxOperations(Val(Me.XPTxtID)) = False Then
        '         Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
        '         Msg = Msg & Chr(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
        '         MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '         Exit Sub
        '     End If
        ' End If
    
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (txtid.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                rs.delete
       
                StrSQL = "Delete From notes Where NoteID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adCmdText
        
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where NOTES_ID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adCmdText
        
                StrSQL = "Delete  TblChecqueBoxContent  where NoteID =" & val(Me.XPTxtID)
                Cn.Execute StrSQL, , adExecuteNoRecords
        
                ' StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text)
                ' Cn.Execute StrSQL, , adExecuteNoRecords

                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    ClearMe
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                    GetDealerInstallment
                End If
            End If
        End If

    Else
        ClearMe
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

Public Sub Retrive(Optional Lngid As Long, Optional NoteID As Long = 0)
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RowNum As Integer
    Dim IntTemp As Integer
    Dim RsDev As ADODB.Recordset
    Dim i As Integer
    Dim NoteSerial1 As String
    On Error GoTo ErrTrap

    If rs.EOF Or rs.BOF Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "ReceiptID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If



    If NoteID <> 0 Then
        rs.find "NoteID=" & NoteID, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If


    txtid.Text = IIf(IsNull(rs("ReceiptID").value), "", rs("ReceiptID").value)
    TXTMessageDES.Text = IIf(IsNull(rs("MessageDES").value), "", rs("MessageDES").value)

    Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)
    Me.CboType.ListIndex = IIf(IsNull(rs("OperationType").value), 0, rs("OperationType").value)
    DBCboClientName.BoundText = IIf(IsNull(rs("Cust_ID").value), "", rs("Cust_ID").value)

    If Me.DBCboClientName.BoundText <> "" Then
        IntTemp = GetDealerType(val(Me.DBCboClientName.BoundText))

        If IntTemp = 1 Then
            Me.CboDealerType.ListIndex = 0
        ElseIf IntTemp = 2 Then
            Me.CboDealerType.ListIndex = 1
        Else
            Me.CboDealerType.ListIndex = -1
        End If
    End If

    DtbBill.value = IIf(IsNull(rs("ReceiptDate").value), "", rs("ReceiptDate").value)
    CboResType.ListIndex = IIf(IsNull(rs("ReceiptType").value), 0, rs("ReceiptType").value)

    If Not IsNull(rs("DiscountType").value) Then
        CboPrecenType.ListIndex = rs("DiscountType").value
    Else
        CboPrecenType.ListIndex = 2
    End If

    Txt(3).Text = IIf(IsNull(rs("DiscounVal").value), "", (rs("DiscounVal").value))

    txtRemark.Text = IIf(IsNull(rs("Remark").value), "", (rs("Remark").value))

    TxtQastNO.Text = IIf(IsNull(rs("PartCount").value), "", rs("PartCount").value)
    TxtSum.Text = IIf(IsNull(rs("Total").value), "", rs("Total").value)
    'Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)

    TXTBankName.Text = IIf(IsNull(rs("BankName").value), "", Trim(rs("BankName").value))

    '-----------------------------------------------------------------------------
    If IsNull(rs("NoteCashingType").value) Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
    
        'project_Expensen_account
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        Me.DCChequeBox.BoundText = ""
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        Me.DCChequeBox.BoundText = ""
    ElseIf rs("NoteCashingType").value = 1 Then
        Me.CboPayMentType.ListIndex = 1
        Me.DcboBox.BoundText = ""
    
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
    
        If SystemOptions.ChequeBox = True Then
            Me.DCChequeBox.BoundText = rs("ChequeBoxID").value
        Else
            Me.DCChequeBox.BoundText = ""
            Me.DcboBankName.BoundText = rs("BankID").value
        End If

    ElseIf rs("NoteCashingType").value = 2 Then

        If SystemOptions.ChequeBox = True Then
            TXTBankName.Visible = True
            'Me.DCChequeBox.BoundText = rs("ChequeBoxID").value
        Else
            TXTBankName.Visible = False
            Me.DCChequeBox.BoundText = ""
            Me.DcboBankName.BoundText = rs("BankID").value
        End If

        Me.CboPayMentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        Me.DCChequeBox.BoundText = ""

    ElseIf rs("NoteCashingType").value = 3 Then

                        If SystemOptions.ChequeBox = True Then
                            TXTBankName.Visible = True
                            'Me.DCChequeBox.BoundText = rs("ChequeBoxID").value
                        Else
                            TXTBankName.Visible = False
                            Me.DCChequeBox.BoundText = ""
                            Me.DcboBankName.BoundText = rs("BankID").value
                        End If

        Me.CboPayMentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.Text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        Me.DCChequeBox.BoundText = ""

   ElseIf rs("NoteCashingType").value = 4 Then
        Me.CboPayMentType.ListIndex = 4
        Me.DCAccounts.BoundText = IIf(IsNull(rs("AccountCode").value), "", rs("AccountCode").value)
        DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.Text = ""
        '    DCVendor.BoundText = ""
     
 
    
    End If

    CboPayMentType_Change

    TxtValue.Text = IIf(IsNull(rs("PaymentMoney").value), "", rs("PaymentMoney").value)
    DCboUserName.BoundText = IIf(IsNull(rs("User_ID").value), "", rs("User_ID").value)
    Txt_akchen.Text = IIf(IsNull(rs("NumAkch").value), "", rs("NumAkch").value)

    'Rs("NumAkch").Value = iff(Txt_akchen.text = "", "", Txt_akchen.text)
    XPTxtID.Text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
    TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)

    Me.oldtxtNoteSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(30).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Dcbranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)

    FgDetails.Rows = 2
    FgDetails.Clear flexClearScrollable, flexClearEverything

    If txtid.Text <> "" Then
        StrSQL = "select * From Qest_Had_Receipted where ReceiptID=" & txtid.Text
        '    StrSql = StrSql + " and  Status <>1"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then

            With FgDetails
                .Rows = RsTemp.RecordCount + 1

                For RowNum = 1 To RsTemp.RecordCount
                    .Rowdata(RowNum) = IIf(IsNull(RsTemp("QestID").value), "", (RsTemp("QestID").value))
                    .TextMatrix(RowNum, .ColIndex("Serial")) = IIf(IsNull(RsTemp("QeqtNum").value), "", RsTemp("QeqtNum").value)
                    .TextMatrix(RowNum, .ColIndex("BillID")) = IIf(IsNull(RsTemp("Transaction_Serial").value), "", RsTemp("Transaction_Serial").value)
                   
                    GetTransNoteSerial1TransactionSerail (.TextMatrix(RowNum, .ColIndex("BillID"))), NoteSerial1
                    .TextMatrix(RowNum, .ColIndex("NoteSerial1")) = NoteSerial1
                
                    .TextMatrix(RowNum, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", RsTemp("Value").value)

                    If Not IsNull(RsTemp("ReceiptDate").value) Then
                        .TextMatrix(RowNum, .ColIndex("Due_Date")) = DisplayDate(RsTemp("ReceiptDate").value)
                    End If

                    RsTemp.MoveNext
                Next RowNum

            End With

        End If
    End If

    '----------------------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then
        Me.DcboDebitSide.BoundText = ""
        Me.DcboDebitSide1.BoundText = ""
        Me.DcboCreditSide.BoundText = ""
        Me.DcboCreditSide1.BoundText = ""
        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where ReceiptID=" & val(Me.txtid.Text)
        StrSQL = StrSQL + " AND Credit_Or_Debit=0"
        StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lbl(33).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For i = 1 To RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                    If RsDev("DEV_ID_Line_No").value = 1 Then
                        Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                    Else
                        Me.DcboDebitSide1.BoundText = RsDev("Account_Code").value
                    End If
                End If

                RsDev.MoveNext
            Next i

        End If

        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where ReceiptID=" & val(Me.txtid.Text)
        StrSQL = StrSQL + " AND Credit_Or_Debit=1"
        StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            If RsDev.RecordCount = 1 Then
                Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
            Else
                Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                RsDev.MoveNext
                Me.DcboCreditSide1.BoundText = RsDev("Account_Code").value
            End If
        End If
    End If

    '----------------------------------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub ClearMe()
    On Error GoTo ErrTrap
    clear_all Me
    txtid.Text = CStr(new_id("ReceiptQest", "ReceiptID", "", True))
    Me.DCboUserName.BoundText = user_id
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FgDetails.Rows = 2
    FgDetails.Clear flexClearScrollable, flexClearEverything
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadIcons()
    On Error GoTo ErrTrap

    '«·√ﬁ”«ÿ «·„ƒÃ·…
    With FG
        .Cell(flexcpPicture, 0, .ColIndex("Serial")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillID")) = mdifrmmain.ImgLstTree.ListImages("Sall").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Value")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Due_Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Released")) = mdifrmmain.ImgLstTree.ListImages("Tick").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
        .AutoSize 0, .Cols - 1, False
    End With

    ''«·√ﬁ”«ÿ «·„Õ’·…
    With FgReceipted
        .Cell(flexcpPicture, 0, .ColIndex("Serial")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillID")) = mdifrmmain.ImgLstTree.ListImages("Sall").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Code")) = mdifrmmain.ImgLstTree.ListImages("code").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Value")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Due_Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
        .AutoSize 0, .Cols - 1, False
    End With

    ''«·√ﬁ”«ÿ «·„”œœ… ›Ì «·⁄„·Ì… «·Õ«·Ì…
    With FgDetails
        .Cell(flexcpPicture, 0, .ColIndex("Serial")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillID")) = mdifrmmain.ImgLstTree.ListImages("Sall").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Value")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Due_Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
        .AutoSize 0, .Cols - 1, False
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
ErrTrap:         End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
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

    For i = LBound(cDcboSearch) To UBound(cDcboSearch)
        Set cDcboSearch(i) = Nothing
    Next i

    Exit Sub
ErrTrap:
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

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hWnd, " Õ’Ì· «·√ﬁ”«ÿ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "· ”ÃÌ· ⁄„·Ì…  Õ’Ì· ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· «·√ﬁ”«ÿ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ⁄„·Ì… «· Õ’Ì· «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· «·√ﬁ”«ÿ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «· Õ’Ì· «·Õ«·Ì…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· «·√ﬁ”«ÿ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· «·√ﬁ”«ÿ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· «·√ﬁ”«ÿ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· «·√ﬁ”«ÿ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· «·√ﬁ”«ÿ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· «·√ﬁ”«ÿ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, " Õ’Ì· «·√ﬁ”«ÿ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub GetDealerInstallment()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RowNum As Integer
    Dim Msg As String
    Dim NoteSerial1 As String
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    Me.lbl(12).Caption = ""
    Me.lbl(13).Caption = ""

    If Me.CboType.ListIndex = -1 Then
        Exit Sub
    End If

    If DBCboClientName.BoundText <> "" Then
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "select * From QryCust_Qest where CustID=" & DBCboClientName.BoundText
            StrSQL = StrSQL + " and QestID not in(select QestID From InstallmentDet_Junc_Receipt where Status<>1)"

            If Me.CboType.ListIndex = 0 Then
                StrSQL = StrSQL + " and QryCust_Qest.Type=0"
            ElseIf Me.CboType.ListIndex = 1 Then
                StrSQL = StrSQL + " and QryCust_Qest.Type=1"
            End If

            '—«Ã⁄ Â–Â «·‰ﬁÿ… ﬂÊÌ” „⁄ «·‹ SQL Server
            StrSQL = StrSQL + " And (ISNULL(QryCust_Qest.Summition) OR  (QryCust_Qest.Summition  <>  [QryCust_Qest].[Value]))"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = " Select QryCust_Qest.* From QryCust_Qest"
            StrSQL = StrSQL + " Where CustID =" & DBCboClientName.BoundText & ""
            StrSQL = StrSQL + " And QestID NOT in(select QestID From InstallmentDet_Junc_Receipt where Status<>1)"

            If Me.CboType.ListIndex = 0 Then
                StrSQL = StrSQL + " and QryCust_Qest.Type=0"
            ElseIf Me.CboType.ListIndex = 1 Then
                StrSQL = StrSQL + " and QryCust_Qest.Type=1"
            End If

            StrSQL = StrSQL + " And ((QryCust_Qest.Summition IS NULL) OR  (QryCust_Qest.Summition  <>  [QryCust_Qest].[Value]))"
        End If

        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        FG.Rows = FG.FixedRows

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            Me.Img.Visible = False
            Msg = "«·√ﬁ”«ÿ «·„”Ã·… "

            If Me.CboType.ListIndex = 0 Then
                Msg = Msg + "··‘—ﬂ… ⁄·Ï «·⁄„Ì· " & Me.DBCboClientName.Text
            ElseIf Me.CboType.ListIndex = 1 Then
                Msg = Msg + "··⁄„Ì· " & Me.DBCboClientName.Text & " ⁄·Ï «·‘—ﬂ… "
            End If

            FG.Rows = RsTemp.RecordCount + 1

            For RowNum = 1 To RsTemp.RecordCount

                With FG
                    .TextMatrix(RowNum, .ColIndex("Serial")) = RowNum
                    .Rowdata(RowNum) = IIf(IsNull(RsTemp("QestID").value), "", (RsTemp("QestID").value))
                    .TextMatrix(RowNum, .ColIndex("QeqtNum")) = IIf(IsNull(RsTemp("QeqtNum").value), "", (RsTemp("QeqtNum").value))
                    .TextMatrix(RowNum, .ColIndex("BillID")) = IIf(IsNull(RsTemp("Transaction_ID").value), "", (RsTemp("Transaction_ID").value))
                
                    GetTransNoteSerial1FromID val(.TextMatrix(RowNum, .ColIndex("BillID"))), NoteSerial1
                    .TextMatrix(RowNum, .ColIndex("NoteSerial1")) = NoteSerial1
                
                    .TextMatrix(RowNum, .ColIndex("Transaction_Serial")) = IIf(IsNull(RsTemp("Transaction_Serial").value), "", (RsTemp("Transaction_Serial").value))
                    .TextMatrix(RowNum, .ColIndex("TransactionTypeName")) = IIf(IsNull(RsTemp("TransactionTypeName").value), "", (RsTemp("TransactionTypeName").value))
                    .TextMatrix(RowNum, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", (RsTemp("Value").value))
                    .TextMatrix(RowNum, .ColIndex("Due_Date")) = IIf(IsNull(RsTemp("DueDate").value), "", Format(RsTemp("DueDate").value, "yyyy/mm/dd"))
                    .TextMatrix(RowNum, .ColIndex("Des")) = IIf(IsNull(RsTemp("Summition").value), "·« ÌÊÃœ", (RsTemp("Summition").value))

                    If RsTemp("Summition").value <> "" Then
                        .Cell(flexcpBackColor, RowNum, 0, RowNum, .Cols - 1) = vbGreen
                    End If

                End With

                RsTemp.MoveNext
            Next RowNum

        Else
            Msg = "·« ÊÃœ «Ì… √ﬁ”«ÿ "

            If Me.CboType.ListIndex = 0 Then
                Msg = Msg + "··‘—ﬂ… ⁄·Ï «·⁄„Ì· " & Me.DBCboClientName.Text
            ElseIf Me.CboType.ListIndex = 1 Then
                Msg = Msg + "··⁄„Ì· " & Me.DBCboClientName.Text & " ⁄·Ï «·‘—ﬂ… "
            End If
        
            Me.Img.Visible = True
        End If
    End If

    Me.lbl(12).Caption = Msg
    '-----------------------------------------
    '«·√ﬁ”«ÿ «· Ï œ›⁄  «Ê Õ’·  ·Â–« «·⁄„Ì· ﬁ»· –·ﬂ
    FgReceipted.Rows = FgReceipted.FixedRows
    FgReceipted.Clear flexClearScrollable, flexClearEverything

    FgPayed.Rows = FgPayed.FixedRows
    FgPayed.Clear flexClearScrollable, flexClearEverything

    If DBCboClientName.BoundText <> "" Then
        StrSQL = "select * From Qest_Had_Receipted where CustID=" & DBCboClientName.BoundText
        StrSQL = StrSQL + " AND Qest_Had_Receipted.Type=0"
        StrSQL = StrSQL + " Order By Qest_Had_Receipted.ReceiptID,Qest_Had_Receipted.QestID"
    
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then

            With FgReceipted
                .Rows = RsTemp.RecordCount + 1

                For RowNum = 1 To RsTemp.RecordCount
                    .TextMatrix(RowNum, .ColIndex("Serial")) = RowNum
                    .TextMatrix(RowNum, .ColIndex("QeqtNum")) = IIf(IsNull(RsTemp("QeqtNum").value), "", RsTemp("QeqtNum").value)
                    .TextMatrix(RowNum, .ColIndex("BillID")) = IIf(IsNull(RsTemp("Transaction_ID").value), "", RsTemp("Transaction_ID").value)
                
                    GetTransNoteSerial1FromID val(.TextMatrix(RowNum, .ColIndex("BillID"))), NoteSerial1
                    .TextMatrix(RowNum, .ColIndex("NoteSerial1")) = NoteSerial1
                
                    .TextMatrix(RowNum, .ColIndex("TransactionTypeName")) = IIf(IsNull(RsTemp("TransactionTypeName").value), "", RsTemp("TransactionTypeName").value)
                    .TextMatrix(RowNum, .ColIndex("Transaction_Serial")) = IIf(IsNull(RsTemp("Transaction_Serial").value), "", (RsTemp("Transaction_Serial").value))
                    .TextMatrix(RowNum, .ColIndex("Code")) = IIf(IsNull(RsTemp("ReceiptID").value), "", RsTemp("ReceiptID").value)
                    .TextMatrix(RowNum, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", RsTemp("Value").value)

                    If Not IsNull(RsTemp("ReceiptDate").value) Then
                        .TextMatrix(RowNum, .ColIndex("Due_Date")) = DisplayDate(RsTemp("ReceiptDate").value)
                    End If
                
                    If RsTemp("ReceiptType").value = 0 Then
                        .TextMatrix(RowNum, .ColIndex("ReceiptType")) = " Õ’Ì· ⁄«œÏ"
                    ElseIf RsTemp("ReceiptType").value = 1 Then
                        .TextMatrix(RowNum, .ColIndex("ReceiptType")) = " Õ’Ì· »Œ’„"
                    ElseIf RsTemp("ReceiptType").value = 2 Then
                        .TextMatrix(RowNum, .ColIndex("ReceiptType")) = "œ›⁄… „‰    «·ﬁ”ÿ"
                    ElseIf RsTemp("ReceiptType").value = 3 Then
                        .TextMatrix(RowNum, .ColIndex("ReceiptType")) = "œ›⁄Â „ﬁœ„…"
                    ElseIf RsTemp("ReceiptType").value = 4 Then
                        .TextMatrix(RowNum, .ColIndex("ReceiptType")) = "œ›⁄Â „‰ «·Õ”«»"
                    End If
                
                    RsTemp.MoveNext
                Next RowNum

                .AutoSize 0, .Cols - 1, False
            End With

        End If

        StrSQL = "Select * From Qest_Had_Receipted where CustID=" & DBCboClientName.BoundText
        StrSQL = StrSQL + " AND Qest_Had_Receipted.Type=1"
        StrSQL = StrSQL + " Order By Qest_Had_Receipted.ReceiptID,Qest_Had_Receipted.QestID"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then

            With FgPayed
                .Rows = RsTemp.RecordCount + 1

                For RowNum = 1 To RsTemp.RecordCount
                    .TextMatrix(RowNum, .ColIndex("Serial")) = RowNum
                    .TextMatrix(RowNum, .ColIndex("QeqtNum")) = IIf(IsNull(RsTemp("QeqtNum").value), "", RsTemp("QeqtNum").value)
                    .TextMatrix(RowNum, .ColIndex("BillID")) = IIf(IsNull(RsTemp("Transaction_ID").value), "", RsTemp("Transaction_ID").value)
                
                    GetTransNoteSerial1FromID val(.TextMatrix(RowNum, .ColIndex("BillID"))), NoteSerial1
                    .TextMatrix(RowNum, .ColIndex("NoteSerial1")) = NoteSerial1
                
                    .TextMatrix(RowNum, .ColIndex("TransactionTypeName")) = IIf(IsNull(RsTemp("TransactionTypeName").value), "", RsTemp("TransactionTypeName").value)
                    .TextMatrix(RowNum, .ColIndex("Transaction_Serial")) = IIf(IsNull(RsTemp("Transaction_Serial").value), "", (RsTemp("Transaction_Serial").value))
                    .TextMatrix(RowNum, .ColIndex("Code")) = IIf(IsNull(RsTemp("ReceiptID").value), "", RsTemp("ReceiptID").value)
                    .TextMatrix(RowNum, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "", RsTemp("Value").value)

                    If Not IsNull(RsTemp("ReceiptDate").value) Then
                        .TextMatrix(RowNum, .ColIndex("Due_Date")) = DisplayDate(RsTemp("ReceiptDate").value)
                    End If

                    If RsTemp("ReceiptType").value = 0 Then
                        .TextMatrix(RowNum, .ColIndex("ReceiptType")) = " Õ’Ì· ⁄«œÏ"
                    ElseIf RsTemp("ReceiptType").value = 1 Then
                        .TextMatrix(RowNum, .ColIndex("ReceiptType")) = " Õ’Ì· »Œ’„"
                    ElseIf RsTemp("ReceiptType").value = 2 Then
                        .TextMatrix(RowNum, .ColIndex("ReceiptType")) = "œ›⁄… „‰   «·ﬁ”ÿ"
                
                    ElseIf RsTemp("ReceiptType").value = 3 Then
                        .TextMatrix(RowNum, .ColIndex("ReceiptType")) = "  œ›⁄Â „ﬁœ„Â "
                    ElseIf RsTemp("ReceiptType").value = 4 Then
                        .TextMatrix(RowNum, .ColIndex("ReceiptType")) = "œ›⁄… „‰ «·Õ”«»  "
                
                    End If

                    RsTemp.MoveNext
                Next RowNum

                .AutoSize 0, .Cols - 1, False
            End With

        End If
    End If

    '----------------------------------------------------
    Exit Sub
ErrTrap:
End Sub

Private Sub SetReleaseType()
    FG.Enabled = True
    lbl(16).Visible = False
    CboPrecenType.Visible = False
    lbl(17).Visible = False
    Txt(3).Visible = False

    If Me.CboResType.ListIndex = -1 Then
        Exit Sub
    ElseIf Me.CboResType.ListIndex = 0 Then
        TxtValue.locked = True
        TxtValue.Enabled = False

    ElseIf Me.CboResType.ListIndex = 1 Then
        TxtValue.locked = True
        TxtValue.Enabled = False

        lbl(16).Visible = True
        CboPrecenType.Visible = True
        lbl(17).Visible = True
        Txt(3).Visible = True
    
    ElseIf Me.CboResType.ListIndex = 3 Or CboResType.ListIndex = 4 Then
        FG.Enabled = False
        TxtValue.locked = False
        TxtValue.Enabled = True
        TxtValue.Text = 0
    
        If CalcNoOfInstallments(val(Me.TxtValue)) = False Then
            Exit Sub
        Else
    
        End If
    
    Else
        TxtValue.locked = False
        TxtValue.Enabled = True
    End If

    WriteDev
End Sub

Private Sub WriteDev()
    Dim Account_Code_dynamic As String

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Me.DcboDebitSide1.BoundText = ""
        Me.DcboCreditSide1.BoundText = ""

        If Me.CboType.ListIndex = 0 Then
            ' Õ’Ì· √ﬁ”«ÿ
            '«·ÿ—› «·„œÌ‰
            Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))

            '«· Õ’Ì· »Œ’„
            If Me.CboResType.ListIndex = 1 Then
        
                Account_Code_dynamic = get_account_code_branch(12, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    Exit Sub
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  Œ’„ „”„ÊÕ »Â ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        Exit Sub
         
                    End If
                End If

                Me.DcboDebitSide1.BoundText = Account_Code_dynamic
                ' Me.DcboDebitSide1.BoundText = "a3a5" ' Œ’„ „”„ÊÕ »Â
            Else
                Me.DcboDebitSide1.BoundText = ""
            End If

            '«·ÿ—› «·œ«∆‰
            Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
        ElseIf Me.CboType.ListIndex = 1 Then
            '”œ«œ √ﬁ”«ÿ
            '«·ÿ—› «·„œÌ‰
            Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))

            '«· Õ’Ì· »Œ’„
            If Me.CboResType.ListIndex = 1 Then
                 
                Account_Code_dynamic = get_account_code_branch(13, my_branch)
                 
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    Exit Sub
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  Œ’„ „ﬂ ”» »Â ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        Exit Sub
                  
                    End If
                End If

                Me.DcboCreditSide1.BoundText = Account_Code_dynamic
                'Me.DcboCreditSide1.BoundText = "a4a4" 'Œ’„ „ﬂ ”»
            Else
                Me.DcboCreditSide1.BoundText = ""
            End If

            '«·ÿ—› «·œ«∆‰
            Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
        End If
    End If

End Sub
