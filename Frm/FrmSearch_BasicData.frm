VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmSearch_BasicData 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13260
   Icon            =   "FrmSearch_BasicData.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_confirmVacation 
      BackColor       =   &H00E2E9E9&
      Height          =   8172
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame16 
         BackColor       =   &H00E2E9E9&
         Height          =   1332
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   6720
         Width           =   12972
         Begin VB.TextBox txtIDConfrimVacation 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   9552
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   264
            Width           =   2328
         End
         Begin MSDataListLib.DataCombo dcDurationConfrimVacation 
            Height          =   288
            Left            =   5640
            TabIndex        =   105
            Top             =   264
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMonthConfrimVacation 
            Height          =   288
            Left            =   2160
            TabIndex        =   106
            Top             =   264
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcVacationType 
            Height          =   288
            Left            =   2160
            TabIndex        =   112
            Top             =   720
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCityCV 
            Height          =   288
            Left            =   9552
            TabIndex        =   113
            Top             =   720
            Width           =   2328
            _ExtentX        =   4101
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcMangerialArea 
            Height          =   288
            Left            =   5640
            TabIndex        =   114
            Top             =   720
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄ÿ·…"
            Height          =   312
            Index           =   26
            Left            =   4368
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   720
            Width           =   984
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Õ«›Ÿ…"
            ForeColor       =   &H00000000&
            Height          =   372
            Left            =   11784
            TabIndex        =   116
            Top             =   720
            Width           =   1068
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‰ÿﬁ… «·«œ«—Ì…"
            ForeColor       =   &H00000000&
            Height          =   372
            Left            =   7704
            TabIndex        =   115
            Top             =   720
            Width           =   1068
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   288
            Index           =   20
            Left            =   4536
            TabIndex        =   109
            Top             =   240
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   288
            Index           =   19
            Left            =   7800
            TabIndex        =   108
            Top             =   240
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   288
            Index           =   17
            Left            =   11856
            TabIndex        =   107
            Top             =   264
            Width           =   960
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E2E9E9&
         Height          =   6012
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   720
         Width           =   12972
         Begin VSFlex8UCtl.VSFlexGrid fg_confirmVacation 
            Height          =   5796
            Left            =   0
            TabIndex        =   102
            Top             =   120
            Width           =   12948
            _cx             =   22839
            _cy             =   10223
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmSearch_BasicData.frx":038A
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   600
         Left            =   0
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   0
         Width           =   13260
         _cx             =   23389
         _cy             =   1058
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     »ÕÀ «À»«   ⁄ÿ·       "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
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
         CaptionStyle    =   1
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
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
   End
   Begin VB.Frame frm_StopDealing 
      BackColor       =   &H00E2E9E9&
      Height          =   8172
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   82
      Top             =   0
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame14 
         BackColor       =   &H00E2E9E9&
         Height          =   6012
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   720
         Width           =   12972
         Begin VSFlex8UCtl.VSFlexGrid fg_StopDealing 
            Height          =   5796
            Left            =   0
            TabIndex        =   96
            Top             =   120
            Width           =   12948
            _cx             =   22839
            _cy             =   10223
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
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmSearch_BasicData.frx":0519
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
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E2E9E9&
         Height          =   1212
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   6720
         Width           =   12972
         Begin VB.ComboBox cbViolation 
            Height          =   288
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   240
            Visible         =   0   'False
            Width           =   2412
         End
         Begin VB.TextBox txtIDStopDealing 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   10920
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   264
            Width           =   1332
         End
         Begin VB.TextBox txtCodeStopDealing 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   10920
            MaxLength       =   50
            TabIndex        =   85
            Top             =   600
            Width           =   1332
         End
         Begin VB.TextBox txtRecordNoStopDealing 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   7560
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   600
            Width           =   2040
         End
         Begin MSDataListLib.DataCombo dcIDACStopDealing 
            Height          =   288
            Left            =   7560
            TabIndex        =   87
            Top             =   264
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcVendorStopDealing 
            Height          =   288
            Left            =   3960
            TabIndex        =   88
            Top             =   624
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   288
            Index           =   18
            Left            =   11856
            TabIndex        =   94
            Top             =   264
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄ﬁœ"
            Height          =   288
            Index           =   15
            Left            =   9936
            TabIndex        =   93
            Top             =   264
            Width           =   720
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ ⁄Âœ"
            Height          =   288
            Index           =   14
            Left            =   6336
            TabIndex        =   92
            Top             =   600
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Œ«·›…"
            Height          =   288
            Index           =   13
            Left            =   6312
            TabIndex        =   91
            Top             =   264
            Visible         =   0   'False
            Width           =   924
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂÊœ"
            Height          =   312
            Index           =   12
            Left            =   12408
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   600
            Width           =   372
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã· "
            Height          =   312
            Index           =   11
            Left            =   9948
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   600
            Width           =   768
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   600
         Left            =   0
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   0
         Width           =   13260
         _cx             =   23389
         _cy             =   1058
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     »ÕÀ «Ìﬁ«› «·„⁄œ« /«·”Ì«—«        "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
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
         CaptionStyle    =   1
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
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
   End
   Begin VB.Frame frm_confrimviolation 
      BackColor       =   &H00E2E9E9&
      Height          =   8172
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   0
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame12 
         BackColor       =   &H00E2E9E9&
         Height          =   1212
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   6720
         Width           =   12972
         Begin VB.TextBox txtRecordno 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   7560
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   600
            Width           =   2040
         End
         Begin VB.TextBox txtfullcode 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   10920
            MaxLength       =   50
            TabIndex        =   78
            Top             =   600
            Width           =   1332
         End
         Begin VB.TextBox txtid4 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   10920
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   264
            Width           =   1332
         End
         Begin MSDataListLib.DataCombo DurationID 
            Height          =   288
            Left            =   492
            TabIndex        =   65
            Top             =   264
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo MinistryContractID 
            Height          =   288
            Left            =   7560
            TabIndex        =   66
            Top             =   264
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcvendor4 
            Height          =   288
            Left            =   3960
            TabIndex        =   73
            Top             =   624
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo MonthID 
            Height          =   288
            Left            =   480
            TabIndex        =   76
            Top             =   600
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo ViolationID 
            Height          =   288
            Left            =   3960
            TabIndex        =   77
            Top             =   264
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã· "
            Height          =   312
            Index           =   21
            Left            =   9948
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   600
            Width           =   768
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂÊœ"
            Height          =   312
            Index           =   22
            Left            =   12408
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   600
            Width           =   372
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Œ«·›…"
            Height          =   288
            Index           =   10
            Left            =   6312
            TabIndex        =   75
            Top             =   264
            Width           =   924
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ ⁄Âœ"
            Height          =   288
            Index           =   9
            Left            =   6336
            TabIndex        =   74
            Top             =   600
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄ﬁœ"
            Height          =   288
            Index           =   8
            Left            =   9936
            TabIndex        =   70
            Top             =   264
            Width           =   720
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   288
            Index           =   7
            Left            =   2712
            TabIndex        =   69
            Top             =   624
            Width           =   804
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   288
            Index           =   5
            Left            =   11856
            TabIndex        =   68
            Top             =   264
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   288
            Index           =   3
            Left            =   2616
            TabIndex        =   67
            Top             =   264
            Width           =   960
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E2E9E9&
         Height          =   6012
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   720
         Width           =   12972
         Begin VSFlex8UCtl.VSFlexGrid fg_ConfirmViolation 
            Height          =   5796
            Left            =   0
            TabIndex        =   62
            Top             =   120
            Width           =   12948
            _cx             =   22839
            _cy             =   10223
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
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmSearch_BasicData.frx":0669
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   600
         Left            =   0
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   0
         Width           =   13260
         _cx             =   23389
         _cy             =   1058
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     »ÕÀ «À»«  «·„Œ«·›«        "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
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
         CaptionStyle    =   1
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
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   972
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   8160
      Width           =   13212
      Begin VB.CommandButton cmd 
         Caption         =   "Œ—ÊÃ"
         Height          =   492
         Index           =   2
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   360
         Width           =   972
      End
      Begin VB.CommandButton cmd 
         Caption         =   "„”Õ "
         Height          =   492
         Index           =   1
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   360
         Width           =   972
      End
      Begin VB.CommandButton cmd 
         Caption         =   "»ÕÀ "
         Height          =   492
         Index           =   0
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   360
         Width           =   972
      End
      Begin VB.Label lblResult 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   300
         RightToLeft     =   -1  'True
         TabIndex        =   118
         Top             =   330
         Width           =   4545
      End
   End
   Begin VB.Frame frm_SchooleFile 
      BackColor       =   &H00E2E9E9&
      Height          =   8172
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   6012
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   12972
         Begin VSFlex8UCtl.VSFlexGrid Fg 
            Height          =   5796
            Left            =   0
            TabIndex        =   13
            Top             =   120
            Width           =   12948
            _cx             =   22839
            _cy             =   10223
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmSearch_BasicData.frx":0840
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
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Height          =   1212
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   6720
         Width           =   12972
         Begin VB.TextBox txtNameE 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   2124
         End
         Begin VB.TextBox txtNameA 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   2052
         End
         Begin VB.ComboBox cbType 
            Height          =   288
            Left            =   2040
            TabIndex        =   5
            Top             =   624
            Width           =   2124
         End
         Begin VB.TextBox txtMinistryNo 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   9600
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   264
            Width           =   2052
         End
         Begin MSDataListLib.DataCombo dcManagrialArea 
            Height          =   288
            Left            =   6012
            TabIndex        =   6
            Top             =   624
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCity 
            Height          =   288
            Left            =   9600
            TabIndex        =   16
            Top             =   600
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‰ÿﬁ… «· ⁄·Ì„Ì…"
            Height          =   285
            Index           =   0
            Left            =   8010
            TabIndex        =   17
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·Ê“«—Ï"
            Height          =   285
            Index           =   4
            Left            =   11850
            TabIndex        =   11
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«”„ «·⁄—»Ï"
            ForeColor       =   &H00000000&
            Height          =   336
            Left            =   8016
            TabIndex        =   10
            Top             =   240
            Width           =   1044
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«”„ «·«‰Ã·Ì“Ï"
            ForeColor       =   &H00000000&
            Height          =   336
            Left            =   4440
            TabIndex        =   9
            Top             =   240
            Width           =   1032
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·œ—”…"
            Height          =   288
            Index           =   6
            Left            =   4152
            TabIndex        =   8
            Top             =   624
            Width           =   1284
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‰ÿﬁ…"
            Height          =   288
            Index           =   16
            Left            =   11496
            TabIndex        =   7
            Top             =   624
            Width           =   1200
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   600
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   13260
         _cx             =   23389
         _cy             =   1058
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     »ÕÀ „·› «·„œ«—”       "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
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
         CaptionStyle    =   1
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
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
   End
   Begin VB.Frame frm_Managerialarea 
      BackColor       =   &H00E2E9E9&
      Height          =   8172
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame9 
         BackColor       =   &H00E2E9E9&
         Height          =   6012
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   720
         Width           =   12972
         Begin VSFlex8UCtl.VSFlexGrid fg_managerialarea 
            Height          =   5796
            Left            =   0
            TabIndex        =   54
            Top             =   120
            Width           =   12948
            _cx             =   22839
            _cy             =   10223
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmSearch_BasicData.frx":09BD
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
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E2E9E9&
         Height          =   972
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   6840
         Width           =   12972
         Begin VB.TextBox txtnameE4 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   360
            Width           =   2124
         End
         Begin VB.TextBox txtname4 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   9720
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   360
            Width           =   2052
         End
         Begin MSDataListLib.DataCombo DcboGovernmentID 
            Height          =   288
            Left            =   2640
            TabIndex        =   49
            Top             =   360
            Width           =   2148
            _ExtentX        =   3784
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«”„ «·⁄—»Ï"
            ForeColor       =   &H00000000&
            Height          =   336
            Left            =   11736
            TabIndex        =   52
            Top             =   360
            Width           =   1044
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«”„ «·«‰Ã·Ì“Ï"
            ForeColor       =   &H00000000&
            Height          =   336
            Left            =   8280
            TabIndex        =   51
            Top             =   360
            Width           =   1032
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ì »⁄ «·Ï"
            ForeColor       =   &H00000000&
            Height          =   372
            Left            =   4932
            TabIndex        =   50
            Top             =   360
            Width           =   696
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   600
         Left            =   0
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   0
         Width           =   13260
         _cx             =   23389
         _cy             =   1058
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     »ÕÀ «·„‰«ÿﬁ «·«œ«—Ì…       "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
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
         CaptionStyle    =   1
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
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
   End
   Begin VB.Frame frm_ViolationTypes 
      BackColor       =   &H00E2E9E9&
      Height          =   8172
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Height          =   6012
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   720
         Width           =   12972
         Begin VSFlex8UCtl.VSFlexGrid fg_ViolationTypes 
            Height          =   5796
            Left            =   0
            TabIndex        =   40
            Top             =   120
            Width           =   12948
            _cx             =   22839
            _cy             =   10223
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
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmSearch_BasicData.frx":0A7E
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
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2E9E9&
         Height          =   972
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   6840
         Width           =   12972
         Begin VB.ComboBox cbDeduct 
            Height          =   288
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   240
            Width           =   2052
         End
         Begin VB.TextBox txtNameE3 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   2124
         End
         Begin VB.TextBox txtName3 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   2052
         End
         Begin VB.TextBox txtcode3 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   1452
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·Œ’„"
            ForeColor       =   &H00000000&
            Height          =   336
            Left            =   2400
            TabIndex        =   43
            Top             =   240
            Width           =   792
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂÊœ"
            Height          =   288
            Index           =   1
            Left            =   11856
            TabIndex        =   38
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«”„ «·⁄—»Ï"
            ForeColor       =   &H00000000&
            Height          =   336
            Left            =   9216
            TabIndex        =   37
            Top             =   240
            Width           =   1044
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«”„ «·«‰Ã·Ì“Ï"
            ForeColor       =   &H00000000&
            Height          =   336
            Left            =   5760
            TabIndex        =   36
            Top             =   240
            Width           =   1032
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   600
         Left            =   0
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Width           =   13260
         _cx             =   23389
         _cy             =   1058
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     »ÕÀ «‰Ê«⁄ «·„Œ«·›«        "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
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
         CaptionStyle    =   1
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
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
   End
   Begin VB.Frame frm_VacationTypes 
      BackColor       =   &H00E2E9E9&
      Height          =   8172
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Height          =   972
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   6840
         Width           =   12972
         Begin VB.TextBox txtcode 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   9600
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   264
            Width           =   2052
         End
         Begin VB.TextBox txtname1 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   2052
         End
         Begin VB.TextBox txtnameE2 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   2124
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«”„ «·«‰Ã·Ì“Ï"
            ForeColor       =   &H00000000&
            Height          =   336
            Left            =   4440
            TabIndex        =   28
            Top             =   240
            Width           =   1032
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«”„ «·⁄—»Ï"
            ForeColor       =   &H00000000&
            Height          =   336
            Left            =   8016
            TabIndex        =   27
            Top             =   240
            Width           =   1044
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂÊœ"
            Height          =   288
            Index           =   2
            Left            =   11856
            TabIndex        =   26
            Top             =   264
            Width           =   840
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   6012
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   12972
         Begin VSFlex8UCtl.VSFlexGrid fg_vacationtype 
            Height          =   5796
            Left            =   0
            TabIndex        =   21
            Top             =   120
            Width           =   12948
            _cx             =   22839
            _cy             =   10223
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmSearch_BasicData.frx":0B87
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   600
         Left            =   0
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   0
         Width           =   13260
         _cx             =   23389
         _cy             =   1058
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     »ÕÀ «‰Ê«⁄ «·⁄ÿ·«        "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
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
         CaptionStyle    =   1
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
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "FrmSearch_BasicData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public calltype As Integer
Public SendForm As String

Private Sub Cmd_Click(Index As Integer)
On Error Resume Next
    lblResult.Caption = ""
    Select Case Index

        Case 0
                If frm_SchooleFile.Visible = True Then
                        GetData
                ElseIf frm_VacationTypes.Visible = True Then
                        GetData_VacationTypes
                ElseIf frm_ViolationTypes.Visible = True Then
                        GetData_ViolationTypes
                ElseIf frm_Managerialarea.Visible = True Then
                        GetData_ManagerialArea
                ElseIf frm_confrimviolation.Visible = True Then
                        GetData_ConfirmViolation
                ElseIf frm_StopDealing.Visible = True Then
                        GetData_StopDealing
                ElseIf frm_confirmVacation.Visible = True Then
                        GetData_ConfirmVacation
                End If
                
        Case 1
            clear_all Me
        Case 2
            Unload Me
    End Select

End Sub

Private Sub dcDurationConfrimVacation_Click(Area As Integer)
Dim str As String

 str = "  select id , Name  from TblDurations_Details where did =   " & val(dcDurationConfrimVacation.BoundText)
fill_combo dcMonthConfrimVacation, str
End Sub

Private Sub dcvendor4_Change()
Dim val1, val2, recordno As String, Fullcode As String
If dcvendor4.BoundText = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and cusid = " & dcvendor4.BoundText
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     End If
     txtRecordno.Text = recordno
     txtfullcode.Text = Fullcode
End Sub



Private Sub dcVendorStopDealing_Click(Area As Integer)
Dim val1, val2, recordno As String, Fullcode As String
If dcVendorStopDealing.BoundText = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and cusid = " & dcVendorStopDealing.BoundText
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     End If
     txtRecordNoStopDealing.Text = recordno
     txtCodeStopDealing.Text = Fullcode
End Sub

Private Sub DurationID_Click(Area As Integer)
Dim i As Integer, j As Integer, str As String
    i = val(DurationID.BoundText)
    
    If i > 0 Then
        str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo MonthID, str
    Else
        str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo Me.MonthID, str
    End If
End Sub

Private Sub Fg_Click()
 On Error GoTo ErrTrap
   Dim i As Integer
            i = val(FG.TextMatrix(FG.Row, FG.ColIndex("ID")))
        If SendForm = "schoolfile" Then
              
               FrmSchooleFile.Retrive (i)
        ElseIf SendForm = "vacationtypes" Then
              FrmVactionTypes.Retrive (i)
        ElseIf SendForm = "VA" Then
             FrmVehicleAllocation.dcSchoolFile.BoundText = i
             
       ElseIf SendForm = "AC2" Then
            With FrmAttributionContract.VSFlexGrid1
                        Dim StrSQL As String
                        FrmAttributionContract.SchoolFile_INfo CStr(i), .Row
            End With
            
            ElseIf SendForm = "SSA" Then
                        FrmSuperVisorSchoolAllocation.dcMA.SetFocus
                        FrmSuperVisorSchoolAllocation.dcMA.BoundText = i
                
            ElseIf SendForm = "DriverAllocation" Then
                        'FrmDriverAllocation.dc
            ElseIf SendForm = "report_scene" Then
                        frmReport_Scenes.dcSchool.BoundText = i
                        
            ElseIf SendForm = "ReportScene_School" Then
                    frmReport_Scenes.dcSchool.BoundText = i
        End If


'Unload Me
ErrTrap:

End Sub

Private Sub fg_confirmVacation_Click()
 On Error GoTo ErrTrap
   Dim i As Integer
     i = val(fg_confirmVacation.TextMatrix(fg_confirmVacation.Row, fg_confirmVacation.ColIndex("id")))
        If i > 0 Then
        
                If SendForm = "confirmVacation" Then
                       FrmConfirmVaction.Retrive (i)
                Else
                      
               End If
        End If
'Unload Me
ErrTrap:
End Sub

Private Sub fg_ConfirmViolation_Click()

 On Error GoTo ErrTrap
   Dim i As Integer
     i = val(fg_ConfirmViolation.TextMatrix(fg_ConfirmViolation.Row, fg_ConfirmViolation.ColIndex("id")))
        If i > 0 Then
        
                If SendForm = "ConfirmViolation" Then
                       FrmConfirmViolation.Retrive (i)
                Else
                      
               End If
        End If
'Unload Me
ErrTrap:



End Sub

Private Sub fg_managerialarea_Click()

 On Error GoTo ErrTrap
   Dim i As Integer
     i = val(fg_managerialarea.TextMatrix(fg_managerialarea.Row, fg_managerialarea.ColIndex("id")))
        If i > 0 Then
        
                If SendForm = "MCMA" Then
                       FrmMinistryContract.DCVendor.BoundText = i
                ElseIf SendForm = "scene" Then
                        frmReport_Scenes.DcMangerialArea.BoundText = i
                ElseIf SendForm = "ExchangeRequest" Then
                        FrmExchangeRequest.dcMangerialAreaID.BoundText = i
                Else
                        FrmManagerialArea.Retrive (i)
               End If
        End If
'Unload Me
ErrTrap:
End Sub

Private Sub fg_StopDealing_Click()
 On Error GoTo ErrTrap
   Dim i As Integer
     i = val(fg_StopDealing.TextMatrix(fg_StopDealing.Row, fg_StopDealing.ColIndex("id")))
        If i > 0 Then
                   FrmStopDealing.Retrive (i)
        End If
'Unload Me
ErrTrap:
End Sub

Private Sub fg_vacationtype_Click()
 On Error GoTo ErrTrap
   Dim i As Integer
     i = val(fg_vacationtype.TextMatrix(fg_vacationtype.Row, fg_vacationtype.ColIndex("id")))
        If i > 0 Then
            If SendForm = "ConfirmVacation" Then
                FrmConfirmVaction.dcVacationType.BoundText = i
            Else
               FrmVactionTypes.Retrive (i)
            End If
        End If
'Unload Me
ErrTrap:
End Sub

Private Sub fg_ViolationTypes_Click()
 On Error GoTo ErrTrap
   Dim i As Integer
     i = val(fg_ViolationTypes.TextMatrix(fg_ViolationTypes.Row, fg_ViolationTypes.ColIndex("id")))
        If i > 0 Then
               FrmViolationTypes.Retrive (i)
        End If
'Unload Me
ErrTrap:
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
PutFormOnTop Me.hwnd, True
mdifrmmain.Enabled = False
End Sub

Private Sub Form_Load()
Dim str As String



    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.getCountriesGovernments dcCity
    Dcombos.GetCustomersSuppliers 2, dcVendorStopDealing
      Dcombos.getCountriesGovernments dcCityCV
    
    Dcombos.GetCustomersSuppliers 2, dcvendor4
    fill_combo MinistryContractID, "select idac ,idac from TblAttributionContract "
    fill_combo DurationID, "select id , name from tbldurations"
    fill_combo ViolationID, "  select id , name  from TblViolationTypes "
    fill_combo dcDurationConfrimVacation, "select id , name from tbldurations"
      
     fill_combo dcIDACStopDealing, "select idac ,idac from TblAttributionContract "
     
     
    str = "   select id , name from TblVacationTypes  "
    fill_combo dcVacationType, str
   
    str = "  select id , name from TblManagerialArea "
    fill_combo DcMangerialArea, str
     
     

    If SystemOptions.UserInterface = ArabicInterface Then
             str = "select id , name  from TblManagerialArea "
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
             str = "select id , namee  from TblManagerialArea "
    End If
    fill_combo dcManagrialArea, str
      Dcombos.getCountriesGovernments Me.DcboGovernmentID
   
    
    Set DCboSearch = New clsDCboSearch
    Set GrdBack = New ClsBackGroundPic
    With Me.FG
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
     
    If SendForm = "ReportScene_School" Or SendForm = "report_scene" Or SendForm = "schoolfile" Or SendForm = "VA" Or SendForm = "AC2" Or SendForm = "SSA" Or SendForm = "DriverAllocation" Then
            frm_SchooleFile.Visible = True
    ElseIf SendForm = "vacationtypes" Or SendForm = "ConfirmVacation" Then
            frm_VacationTypes.Visible = True
    ElseIf SendForm = "violationtypes" Then
            frm_ViolationTypes.Visible = True
    ElseIf SendForm = "ExchangeRequest" Or SendForm = "managerialarea" Or SendForm = "MCMA" Or SendForm = "scene" Then
            frm_Managerialarea.Visible = True
    ElseIf SendForm = "ConfirmViolation" Then
            frm_confrimviolation.Visible = True
    ElseIf SendForm = "stopdealing" Then
            frm_StopDealing.Visible = True
    ElseIf SendForm = "confirmVacation" Then
            frm_confirmVacation.Visible = True
    End If
    
    With cbType
       If SystemOptions.UserInterface = EnglishInterface Then
            .Clear
             .AddItem ("Governmental")
             .AddItem ("Domestic School")
            .AddItem ("International")
        Else
             .Clear
            .AddItem ("ÕﬂÊ„Ï")
            .AddItem ("√Â·Ï")
            .AddItem ("«‰ —‰«‘Ê‰«·")
        End If
    End With
    
    With cbDeduct
    If SystemOptions.UserInterface = ArabicInterface Then
        .Clear
        .AddItem ("ﬁÌ„…")
       .AddItem ("‰”»… „‰ «·«Ã— «·ÌÊ„Ï")
       .AddItem ("‰”»… „‰ „Œ’’ «·ÿ«·»")
       .AddItem ("‰Ê⁄ «·Œ’„ »«·ÿ«·»")
       .AddItem ("‰Ê⁄ «·Œ’„ »«·ÌÊ„")
    Else
     .Clear
       .AddItem ("Value")
      .AddItem ("Percent From Day Salary")
      .AddItem ("Percent From Student Custom")
      .AddItem ("")
      .AddItem ("")
    End If
    End With
       
 
    
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    FG.Rows = FG.FixedRows
  
    StrSQL = "  select *  from TblSchooleFile"
    StrSQL = StrSQL & "     where 1 = 1  "
   
   
    If Me.txtMinistryNo.Text <> "" Then
            StrSQL = StrSQL & "   and  ministerNo like  '%" & txtMinistryNo.Text & "%'"
    End If
    
    If Me.txtNameA.Text <> "" Then
            StrSQL = StrSQL & "   and  Name like  '%" & txtNameA.Text & "%'"
    End If
    
   If Me.txtNamee.Text <> "" Then
            StrSQL = StrSQL & "   and  Namee like  '%" & txtNamee.Text & "%'"
    End If
        
    If Me.dcCity.BoundText <> "" Then
            StrSQL = StrSQL & "   and  cityID =  " & (Me.dcCity.BoundText)
    End If
 
    If Me.cbType.ListIndex <> -1 Then
                StrSQL = StrSQL & "   and  SchooleType   =  " & cbType.ListIndex
    End If
     
    If Me.dcManagrialArea.BoundText <> "" Then
            StrSQL = StrSQL & "   and  cityID =  " & Me.dcManagrialArea.BoundText
    End If
       
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        'MsgBox ("·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ")
        Cmd_Click (1)
        lblResult.Caption = ("·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ")
        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        lblResult.Caption = ("·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ")
        Exit Sub
    Else

        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                .TextMatrix(i, .ColIndex("ministerno")) = IIf(IsNull(rs("ministerno").value), "", rs("ministerno").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs("Namee").value), "", rs("Namee").value)
                .TextMatrix(i, .ColIndex("managerid")) = IIf(IsNull(rs("managerid").value), "", rs("managerid").value)
                .TextMatrix(i, .ColIndex("supervisor")) = IIf(IsNull(rs("supervisor").value), "", rs("supervisor").value)
                .TextMatrix(i, .ColIndex("Telephone")) = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
                 rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub

Public Sub GetData_VacationTypes()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    fg_vacationtype.Rows = fg_vacationtype.FixedRows
  
    StrSQL = "  select *  from TblVacationTypes "
    StrSQL = StrSQL & "     where 1 = 1  "
   
   
    If Me.txtcode.Text <> "" Then
            StrSQL = StrSQL & "   and  code =  " & txtcode.Text
    End If
    
    If Me.txtname1.Text <> "" Then
            StrSQL = StrSQL & "   and  Name like  '%" & txtname1.Text & "%'"
    End If
    
   If Me.txtnameE2.Text <> "" Then
            StrSQL = StrSQL & "   and  Namee like  '%" & txtnameE2.Text & "%'"
    End If
        
    
 
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

      '  Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
     '   MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        lblResult.Caption = ("·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ")
        
        Exit Sub
    Else

        With Me.fg_vacationtype
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("code").value), "", rs("code").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs("Namee").value), "", rs("Namee").value)
                .TextMatrix(i, .ColIndex("remarks")) = IIf(IsNull(rs("remarks").value), "", rs("remarks").value)
             
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub


Public Sub GetData_ViolationTypes()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    fg_ViolationTypes.Rows = fg_ViolationTypes.FixedRows
  
    StrSQL = "  select *  from TblViolationTypes "
    StrSQL = StrSQL & "     where 1 = 1  "
   
   
    If Me.txtcode3.Text <> "" Then
            StrSQL = StrSQL & "   and  code =  " & txtcode3.Text
    End If
    
    If Me.txtName3.Text <> "" Then
            StrSQL = StrSQL & "   and  Name like  '%" & txtName3.Text & "%'"
    End If
    
   If Me.txtNameE3.Text <> "" Then
            StrSQL = StrSQL & "   and  Namee like  '%" & txtNameE3.Text & "%'"
    End If
    
    If cbDeduct.ListIndex <> -1 Then
            StrSQL = StrSQL & " and type =  " & cbDeduct.ListIndex
    End If
    
 
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If
       ' Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        lblResult.Caption = ("·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ")
        Exit Sub
    Else

        With Me.fg_ViolationTypes
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("code").value), "", rs("code").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs("Namee").value), "", rs("Namee").value)
              '  .TextMatrix(i, .ColIndex("remarks")) = IIf(IsNull(rs("remarks").value), "", rs("remarks").value)
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("accountName").value), "", rs("accountName").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub

Public Sub GetData_ManagerialArea()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    fg_managerialarea.Rows = fg_managerialarea.FixedRows
  
    StrSQL = "  select *  from tblmanagerialarea  "
    StrSQL = StrSQL & "     where 1 = 1  "
   
   
   ' If Me.txtcode4.text <> "" Then
   '         StrSQL = StrSQL & "   and  code =  " & txtcode3.text
   ' End If
    
    If Me.txtname4.Text <> "" Then
            StrSQL = StrSQL & "   and  Name like  '%" & txtname4.Text & "%'"
    End If
    
   If Me.txtnameE4.Text <> "" Then
            StrSQL = StrSQL & "   and  Namee like  '%" & txtnameE4.Text & "%'"
    End If
    
    If DcboGovernmentID.BoundText <> "" Then
            StrSQL = StrSQL & " and CityID =  " & DcboGovernmentID.BoundText
    End If
    
 
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

      '  Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
     '   MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        lblResult.Caption = ("·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ")
        Exit Sub
    Else

        With Me.fg_managerialarea
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
             '   .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("code").value), "", rs("code").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs("Namee").value), "", rs("Namee").value)
             '   .TextMatrix(i, .ColIndex("remarks")) = IIf(IsNull(rs("remarks").value), "", rs("remarks").value)
             '   .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs("accountName").value), "", rs("accountName").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub


Public Sub GetData_ConfirmViolation()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    fg_ConfirmViolation.Rows = fg_ConfirmViolation.FixedRows
  
   
StrSQL = StrSQL & "    SELECT distinct  dbo.TblConfirmViolation.VendorID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblViolationTypes.Name AS violationName,"
StrSQL = StrSQL & "                   dbo.TblConfirmViolation.ViolationID, dbo.TblConfirmViolation.DurationID, dbo.TblDurations.Name AS DurName, dbo.TblConfirmViolation.MonthID,"
StrSQL = StrSQL & "                    dbo.TblConfirmViolation.CarID, dbo.TblConfirmViolation.MinistryContractID, dbo.TblConfirmViolation.ViolationType, dbo.TblConfirmViolation.MinistryContractValue,"
StrSQL = StrSQL & "                    dbo.TblConfirmViolation.Date, dbo.TblConfirmViolation.DateH, dbo.TblConfirmViolation.Value, dbo.TblConfirmViolation.AbsenceCount, dbo.TblConfirmViolation.ID,"
StrSQL = StrSQL & "                    dbo.TblDurations_Details.Name AS MonthName, dbo.TblCustemers.Fullcode, dbo.TblCustemers.RecordNo, dbo.TblVehicleAllocation_Details.Type,"
StrSQL = StrSQL & "                    dbo.TblVehicleAllocation_Details.CarID AS Expr1, dbo.TblVehicleAllocation_Details.BoardNo "
StrSQL = StrSQL & "    FROM     dbo.TblConfirmViolation INNER JOIN"
StrSQL = StrSQL & "                     dbo.TblViolationTypes ON dbo.TblConfirmViolation.ViolationID = dbo.TblViolationTypes.ID INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblConfirmViolation.VendorID = dbo.TblCustemers.CusID INNER JOIN"
StrSQL = StrSQL & "                     dbo.TblDurations ON dbo.TblConfirmViolation.DurationID = dbo.TblDurations.ID INNER JOIN"
StrSQL = StrSQL & "                    dbo.TblDurations_Details ON dbo.TblConfirmViolation.MonthID = dbo.TblDurations_Details.ID INNER JOIN"
StrSQL = StrSQL & "                     dbo.TblAttributionContract ON dbo.TblConfirmViolation.MinistryContractID = dbo.TblAttributionContract.IDAC INNER JOIN"
StrSQL = StrSQL & "                     dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA AND"
StrSQL = StrSQL & "                     dbo.TblConfirmViolation.CarID = dbo.TblVehicleAllocation_Details.CarID"
StrSQL = StrSQL & "    Where (dbo.TblVehicleAllocation_Details.Type = 3)"

    
    
    
   
   
    If Me.txtid4.Text <> "" Then
            StrSQL = StrSQL & "   and  TblConfirmViolation.id = " & val(txtid4.Text)
    End If
    
   If Me.dcvendor4.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblConfirmViolation.vendorid = " & val(Me.dcvendor4.BoundText)
    End If
    
    If ViolationID.BoundText <> "" Then
            StrSQL = StrSQL & " and ViolationID =  " & val(ViolationID.BoundText)
            
    End If
    
   
   If Me.DurationID.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblConfirmViolation.DurationID = " & val(Me.DurationID.BoundText)
    End If
    
    If MonthID.BoundText <> "" Then
            StrSQL = StrSQL & " and MonthID =  " & val(MonthID.BoundText)
    End If
    
    If MinistryContractID.BoundText <> "" Then
            StrSQL = StrSQL & " and TblConfirmViolation.MinistryContractID =  " & val(MinistryContractID.BoundText)
    End If
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By TblConfirmViolation.ID  "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            'Me.Lbl(10).Caption = "Search Results=0"
        End If

        'Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
      '  MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        lblResult.Caption = ("·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ")
        Exit Sub
    Else
    
    

        With Me.fg_ConfirmViolation
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
             '   .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("code").value), "", rs("code").value)
                .TextMatrix(i, .ColIndex("MinistryContractID")) = IIf(IsNull(rs("MinistryContractID").value), "", rs("MinistryContractID").value)
                .TextMatrix(i, .ColIndex("DurName")) = IIf(IsNull(rs("DurName").value), "", rs("DurName").value)
                .TextMatrix(i, .ColIndex("MonthName")) = IIf(IsNull(rs("MonthName").value), "", rs("MonthName").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                
                           .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)
                .TextMatrix(i, .ColIndex("recordno")) = IIf(IsNull(rs("recordno").value), "", rs("recordno").value)
                
                   .TextMatrix(i, .ColIndex("violationName")) = IIf(IsNull(rs("violationName").value), "", rs("violationName").value)
                .TextMatrix(i, .ColIndex("MonthName")) = IIf(IsNull(rs("MonthName").value), "", rs("MonthName").value)
                .TextMatrix(i, .ColIndex("BoardNo")) = IIf(IsNull(rs("BoardNo").value), "", rs("BoardNo").value)
                
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub


Public Sub GetData_StopDealing()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    fg_StopDealing.Rows = fg_StopDealing.FixedRows
  
   
'StrSQL = StrSQL & " select * from ( "
'StrSQL = StrSQL & "           SELECT dbo.TblStopDealing.ID, dbo.TblAttributionContract.IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
''StrSQL = StrSQL & "           dbo.TblCustemers.RecordNo, dbo.TblStopDealing.StopDate1 AS StopDate, dbo.TblStopDealing.StopDateH1 AS StopDateH, '' AS BoardNo, '⁄ﬁœ' AS SM"
'StrSQL = StrSQL & "           FROM     dbo.TblStopDealing INNER JOIN"
'StrSQL = StrSQL & "          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID1 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'StrSQL = StrSQL & "          dbo.TblAttributionContract ON dbo.TblStopDealing.AttribID = dbo.TblAttributionContract.IDAC"
''StrSQL = StrSQL & "          Where (dbo.TblStopDealing.StopM = 0)"
'StrSQL = StrSQL & "          Union"
'StrSQL = StrSQL & "         SELECT dbo.TblStopDealing.ID, '' AS IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
''StrSQL = StrSQL & "         dbo.TblCustemers.RecordNo,  dbo.TblStopDealing.FromDate AS StopDate, dbo.TblStopDealing.FromDateH AS StopDateH,"
'StrSQL = StrSQL & "         dbo.TblVendorCars.BoardNo, '”Ì«—…' AS SM"
''StrSQL = StrSQL & "         FROM     dbo.TblStopDealing LEFT OUTER JOIN"
'StrSQL = StrSQL & "          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'StrSQL = StrSQL & "           dbo.TblVendorCars ON dbo.TblStopDealing.CarID = dbo.TblVendorCars.ID"
'StrSQL = StrSQL & "          Where (dbo.TblStopDealing.StopM = 1)"
''StrSQL = StrSQL & " ) tbl1  where 1 = 1 "
       
   
StrSQL = StrSQL & "    select * from ("
          
StrSQL = StrSQL & "    SELECT dbo.TblStopDealing.ID, dbo.TblAttributionContract.IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "    dbo.TblCustemers.Fullcode,           dbo.TblCustemers.RecordNo, dbo.TblStopDealing.StopDate1 AS StopDate, dbo.TblStopDealing.StopDateH1 AS StopDateH,"
StrSQL = StrSQL & "    '' AS BoardNo, '⁄ﬁœ' AS SM"
StrSQL = StrSQL & "    FROM     dbo.TblStopDealing INNER JOIN          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID1 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblAttributionContract ON dbo.TblStopDealing.AttribID = dbo.TblAttributionContract.IDAC"
StrSQL = StrSQL & "    Where (dbo.TblStopDealing.StopM = 0)"
StrSQL = StrSQL & "    Union"


StrSQL = StrSQL & "    SELECT dbo.TblStopDealing.ID, '' AS IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
StrSQL = StrSQL & "    dbo.TblCustemers.RecordNo,  dbo.TblStopDealing.FromDate AS StopDate, dbo.TblStopDealing.FromDateH AS StopDateH,         dbo.TblVendorCars.BoardNo, '”Ì«—…' AS SM"
StrSQL = StrSQL & "    FROM     dbo.TblStopDealing LEFT OUTER JOIN          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID = dbo.TblCustemers.CusID"
StrSQL = StrSQL & "    LEFT OUTER JOIN           dbo.TblVendorCars ON dbo.TblStopDealing.CarID = dbo.TblVendorCars.ID"
StrSQL = StrSQL & "    Where (dbo.TblStopDealing.StopM = 1)"

StrSQL = StrSQL & "    Union"

StrSQL = StrSQL & "    SELECT dbo.TblStopDealing.ID, dbo.TblAttributionContract.IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "    dbo.TblCustemers.Fullcode,           dbo.TblCustemers.RecordNo, dbo.TblStopDealing.StopDate2 AS StopDate, dbo.TblStopDealing.StopDateH2 AS StopDateH,"
StrSQL = StrSQL & "    '' AS BoardNo, '⁄ﬁœ' AS SM"
StrSQL = StrSQL & "    FROM     dbo.TblStopDealing INNER JOIN          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID2 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblAttributionContract ON dbo.TblStopDealing.AttribID2 = dbo.TblAttributionContract.IDAC"
StrSQL = StrSQL & "    Where (dbo.TblStopDealing.stp = 1 And StopDealingType = 2)"

StrSQL = StrSQL & "    Union"

StrSQL = StrSQL & "    SELECT dbo.TblStopDealing.ID, dbo.TblAttributionContract.IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "    dbo.TblCustemers.Fullcode,           dbo.TblCustemers.RecordNo, dbo.TblStopDealing.StopDate3 AS StopDate, dbo.TblStopDealing.StopDateH3 AS StopDateH,"
StrSQL = StrSQL & "    '' AS BoardNo, '⁄ﬁœ' AS SM"
StrSQL = StrSQL & "    FROM     dbo.TblStopDealing INNER JOIN          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID3 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblAttributionContract ON dbo.TblStopDealing.AttribID3 = dbo.TblAttributionContract.IDAC"
StrSQL = StrSQL & "    Where (dbo.TblStopDealing.stp = 1 And StopDealingType = 3)"


StrSQL = StrSQL & "    Union"

StrSQL = StrSQL & "    SELECT dbo.TblStopDealing.ID, dbo.TblAttributionContract.IDAC, dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "    dbo.TblCustemers.Fullcode,           dbo.TblCustemers.RecordNo, dbo.TblStopDealing.StopDate4 AS StopDate, dbo.TblStopDealing.StopDateH4 AS StopDateH,"
StrSQL = StrSQL & "    '' AS BoardNo, '⁄ﬁœ' AS SM"
StrSQL = StrSQL & "    FROM     dbo.TblStopDealing INNER JOIN          dbo.TblCustemers ON dbo.TblStopDealing.CustomerID4 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblAttributionContract ON dbo.TblStopDealing.AttribID4 = dbo.TblAttributionContract.IDAC"
StrSQL = StrSQL & "    Where (dbo.TblStopDealing.stp = 1 And StopDealingType = 4)"

StrSQL = StrSQL & "    ) tb1 where 1= 1"
       
       
       
    If Me.txtIDStopDealing.Text <> "" Then
            StrSQL = StrSQL & "   and  ID= " & val(txtIDStopDealing.Text)
    End If
    
   If Me.dcVendorStopDealing.BoundText <> "" Then
            StrSQL = StrSQL & "   and  cusid = " & val(Me.dcVendorStopDealing.BoundText)
    End If
    
    If dcIDACStopDealing.BoundText <> "" Then
            StrSQL = StrSQL & " and IDAC =  " & val(dcIDACStopDealing.BoundText)
            
    End If
    
   
       
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By ID  "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            'Me.Lbl(10).Caption = "Search Results=0"
        End If

        'Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
       ' MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       lblResult.Caption = ("·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ")
        Exit Sub
    Else
    
    
Dim StopM As Integer
        With Me.fg_StopDealing
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
             
            .TextMatrix(i, .ColIndex("StopM")) = IIf(IsNull(rs("SM").value), "", rs("SM").value)
            .TextMatrix(i, .ColIndex("IDAC")) = IIf(IsNull(rs("IDAC").value), "", rs("IDAC").value)
            .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
            .TextMatrix(i, .ColIndex("recordno")) = IIf(IsNull(rs("recordno").value), "", rs("recordno").value)
            .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
            .TextMatrix(i, .ColIndex("BoardNo")) = IIf(IsNull(rs("BoardNo").value), "", rs("BoardNo").value)
            .TextMatrix(i, .ColIndex("fullcode")) = IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)
            
                
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub

Public Sub GetData_ConfirmVacation()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    fg_StopDealing.Rows = fg_StopDealing.FixedRows
  
   
        StrSQL = StrSQL & "   SELECT dbo.TblCountriesGovernments.GovernmentName, dbo.TblConfirmVacation.CityID, dbo.TblConfirmVacation.DurationID, dbo.TblConfirmVacation.VacationTypeID,"
        StrSQL = StrSQL & "   dbo.TblConfirmVacation.Remarks, dbo.TblConfirmVacation.FromDate, dbo.TblConfirmVacation.ToDate, dbo.TblConfirmVacation.FromDateH,"
        StrSQL = StrSQL & "   dbo.TblConfirmVacation.ToDateH, dbo.TblConfirmVacation.UserID, dbo.TblConfirmVacation.MonthID, dbo.TblConfirmVacation.DayValue,"
        StrSQL = StrSQL & "   dbo.TblConfirmVacation.MangerialAreaID, dbo.TblManagerialArea.Name AS MAName, dbo.TblDurations.Name AS DurName, dbo.TblDurations_Details.Name AS MonthName,"
        StrSQL = StrSQL & "   dbo.TblConfirmVacation.ID, dbo.TblVacationTypes.Name AS TypeName"
        StrSQL = StrSQL & "   FROM     dbo.TblConfirmVacation INNER JOIN"
        StrSQL = StrSQL & "   dbo.TblDurations ON dbo.TblConfirmVacation.DurationID = dbo.TblDurations.ID INNER JOIN"
        StrSQL = StrSQL & "   dbo.TblDurations_Details ON dbo.TblConfirmVacation.MonthID = dbo.TblDurations_Details.ID INNER JOIN"
        StrSQL = StrSQL & "   dbo.TblManagerialArea ON dbo.TblConfirmVacation.MangerialAreaID = dbo.TblManagerialArea.ID INNER JOIN"
        StrSQL = StrSQL & "   dbo.TblCountriesGovernments ON dbo.TblConfirmVacation.CityID = dbo.TblCountriesGovernments.GovernmentID INNER JOIN"
        StrSQL = StrSQL & "   dbo.TblVacationTypes ON dbo.TblConfirmVacation.VacationTypeID = dbo.TblVacationTypes.ID"
        StrSQL = StrSQL & "   Where 1 = 1"
   
    If Me.txtIDConfrimVacation.Text <> "" Then
            StrSQL = StrSQL & "   and  TblConfirmVacation.ID= " & val(txtIDConfrimVacation.Text)
    End If
    
   If Me.dcDurationConfrimVacation.BoundText <> "" Then
            StrSQL = StrSQL & "   and   dbo.TblConfirmVacation.DurationID = " & val(Me.dcDurationConfrimVacation.BoundText)
    End If
    
    If dcMonthConfrimVacation.BoundText <> "" Then
            StrSQL = StrSQL & " and TblConfirmVacation.MonthID =  " & val(dcMonthConfrimVacation.BoundText)
    End If
    
       If dcCityCV.BoundText <> "" Then
            StrSQL = StrSQL & " and TblConfirmVacation.CityID =  " & val(dcCityCV.BoundText)
    End If
    
        If DcMangerialArea.BoundText <> "" Then
            StrSQL = StrSQL & " and TblConfirmVacation.MangerialAreaID =  " & val(DcMangerialArea.BoundText)
    End If
    
        If dcVacationType.BoundText <> "" Then
            StrSQL = StrSQL & " and   TblConfirmVacation.VacationTypeID =  " & val(dcVacationType.BoundText)
    End If
    
    
    
    
          
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By TblConfirmVacation.ID  "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            'Me.Lbl(10).Caption = "Search Results=0"
        End If

     '   Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
       ' MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
       lblResult.Caption = ("·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ")
        Exit Sub
    Else
    
    
Dim StopM As Integer
        With Me.fg_confirmVacation
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
             
            .TextMatrix(i, .ColIndex("TypeName")) = IIf(IsNull(rs("TypeName").value), "", rs("TypeName").value)
            .TextMatrix(i, .ColIndex("DurName")) = IIf(IsNull(rs("DurName").value), "", rs("DurName").value)
            .TextMatrix(i, .ColIndex("MonthName")) = IIf(IsNull(rs("MonthName").value), "", rs("MonthName").value)
            .TextMatrix(i, .ColIndex("GovernmentName")) = IIf(IsNull(rs("GovernmentName").value), "", rs("GovernmentName").value)
            
            .TextMatrix(i, .ColIndex("MAName")) = IIf(IsNull(rs("MAName").value), "", rs("MAName").value)
            .TextMatrix(i, .ColIndex("FromDateH")) = IIf(IsNull(rs("FromDateH").value), "", rs("FromDateH").value)
            .TextMatrix(i, .ColIndex("ToDateH")) = IIf(IsNull(rs("ToDateH").value), "", rs("ToDateH").value)
            .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
                
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub



Private Sub ChangeLang()
 
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
   ' KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub PushButton1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
mdifrmmain.Enabled = True
End Sub

Private Sub txtCodeStopDealing_Change()
Dim val1, val2
If txtCodeStopDealing.Text = "" Then Exit Sub
Dim str As String, recordno As String, CusID As String
recordno = ""
CusID = ""

    str = " select * From TblCustemers where Type=2  and fullcode = '" & txtCodeStopDealing & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
     Else
        txtRecordNoStopDealing.Text = ""
        dcVendorStopDealing.BoundText = ""
    End If
    
    txtRecordNoStopDealing.Text = recordno
    dcVendorStopDealing.BoundText = CusID
End Sub

Private Sub txtfullcode_Change()
Dim val1, val2
If txtfullcode.Text = "" Then Exit Sub
Dim str As String, recordno As String, CusID As String
recordno = ""
CusID = ""

    str = " select * From TblCustemers where Type=2  and fullcode = '" & txtfullcode & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
     Else
        txtRecordno.Text = ""
        dcvendor4.BoundText = ""
    End If
    
    txtRecordno.Text = recordno
    dcvendor4.BoundText = CusID
End Sub


Private Sub txtRecordNo_Change()
Dim val1, val2, CusID As String, Fullcode As String
If txtRecordno.Text = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and recordno = '" & txtRecordno.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
         CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     Else
        dcvendor4.BoundText = ""
        txtfullcode.Text = ""
    End If
    
   dcvendor4.BoundText = CusID
   txtfullcode.Text = Fullcode
End Sub

Private Sub txtRecordNoStopDealing_Change()

Dim val1, val2, CusID As String, Fullcode As String
If txtRecordNoStopDealing.Text = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and recordno = '" & txtRecordNoStopDealing.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
         CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     Else
        dcVendorStopDealing.BoundText = ""
        txtCodeStopDealing.Text = ""
    End If
    
   dcVendorStopDealing.BoundText = CusID
   txtCodeStopDealing.Text = Fullcode


End Sub
