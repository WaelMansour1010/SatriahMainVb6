VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmAnalysItems 
   BackColor       =   &H00E2E9E9&
   Caption         =   " "
   ClientHeight    =   12570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAnalysItems.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12570
   ScaleWidth      =   21765
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   12570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   21765
      _cx             =   38391
      _cy             =   22172
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   12648447
      ForeColor       =   128
      FrontTabColor   =   14871017
      BackTabColor    =   8454143
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   " Õ·Ì·Ì «·«’‰«›| ﬁ«—Ì— «·‘»ﬂ…|ÿ—ﬁ «·œ›⁄|«·›« Ê—… «·«·ﬂ —Ê‰Ì…|New Tab"
      Align           =   5
      CurrTab         =   3
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
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   12150
         Index           =   3
         Left            =   45
         TabIndex        =   126
         Top             =   45
         Width           =   21675
         Begin VB.CheckBox chkNotes 
            Caption         =   "«‘⁄«—« "
            Height          =   345
            Left            =   5520
            TabIndex        =   208
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Export To Excel"
            Height          =   390
            Index           =   1
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   206
            Top             =   5370
            Width           =   2805
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Export To Excel"
            Height          =   390
            Index           =   0
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   300
            Width           =   1695
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete Line"
            Height          =   420
            Left            =   14310
            RightToLeft     =   -1  'True
            TabIndex        =   200
            Top             =   660
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Select All"
            Height          =   255
            Left            =   210
            MaskColor       =   &H00E0E0E0&
            TabIndex        =   198
            Top             =   810
            Width           =   1185
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Warnings only"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   7530
            TabIndex        =   197
            Top             =   5100
            Width           =   2385
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Display the successfully sent invoices"
            Height          =   390
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   4980
            Width           =   2805
         End
         Begin VB.PictureBox Picture1 
            Height          =   2145
            Left            =   18210
            ScaleHeight     =   2085
            ScaleWidth      =   2025
            TabIndex        =   195
            Top             =   1860
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   495
            Left            =   13950
            TabIndex        =   194
            Top             =   60
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Send"
            Height          =   390
            Left            =   12990
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   690
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox Text3 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1410
            PasswordChar    =   "*"
            TabIndex        =   171
            Top             =   180
            Width           =   1215
         End
         Begin VB.CommandButton cmdInsert 
            Caption         =   "Insert"
            Height          =   390
            Left            =   11640
            RightToLeft     =   -1  'True
            TabIndex        =   170
            Top             =   690
            Width           =   1245
         End
         Begin VB.Frame Frame9 
            Height          =   165
            Left            =   360
            TabIndex        =   167
            Top             =   12870
            Width           =   20685
            Begin XtremeSuiteControls.RadioButton RadioButton1 
               Height          =   255
               Index           =   0
               Left            =   13680
               TabIndex        =   168
               Top             =   5880
               Visible         =   0   'False
               Width           =   1335
               _Version        =   786432
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«” ·«„"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RadioButton1 
               Height          =   255
               Index           =   1
               Left            =   15360
               TabIndex        =   169
               Top             =   5520
               Visible         =   0   'False
               Width           =   1335
               _Version        =   786432
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«⁄ „«œ"
               UseVisualStyle  =   -1  'True
            End
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   345
            Left            =   9960
            TabIndex        =   173
            TabStop         =   0   'False
            Top             =   720
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   212729859
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   345
            Left            =   7890
            TabIndex        =   174
            TabStop         =   0   'False
            Top             =   720
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   212729859
            CurrentDate     =   37140
         End
         Begin VSFlex8Ctl.VSFlexGrid grd 
            Height          =   3060
            Index           =   0
            Left            =   390
            TabIndex        =   175
            Top             =   1260
            Width           =   20700
            _cx             =   36512
            _cy             =   5397
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   86
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmAnalysItems.frx":038A
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
            RightToLeft     =   0   'False
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
         Begin VSFlex8Ctl.VSFlexGrid grd 
            Height          =   6240
            Index           =   2
            Left            =   30
            TabIndex        =   176
            Top             =   5730
            Width           =   21000
            _cx             =   37042
            _cy             =   11007
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   71
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmAnalysItems.frx":12CA
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
            RightToLeft     =   0   'False
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
         Begin MSComCtl2.DTPicker ToDate10 
            Height          =   345
            Left            =   5640
            TabIndex        =   177
            TabStop         =   0   'False
            Top             =   5010
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   212729859
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker FromDate10 
            Height          =   345
            Left            =   3660
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   5010
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   212729859
            CurrentDate     =   37140
         End
         Begin ALLButtonS.ALLButton btnNew 
            Height          =   285
            Index           =   2
            Left            =   2700
            TabIndex        =   179
            Top             =   180
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   503
            BTYPE           =   3
            TX              =   "Excel"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16777088
            BCOLO           =   16777088
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmAnalysItems.frx":1FB3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid grd 
            Height          =   2820
            Index           =   1
            Left            =   10410
            TabIndex        =   193
            Top             =   6600
            Visible         =   0   'False
            Width           =   9660
            _cx             =   17039
            _cy             =   4974
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmAnalysItems.frx":1FCF
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   495
            Left            =   20190
            TabIndex        =   199
            TabStop         =   0   'False
            Top             =   5280
            Width           =   1095
            _cx             =   1931
            _cy             =   873
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
         End
         Begin MSDataListLib.DataCombo DcBranches 
            Height          =   315
            Index           =   2
            Left            =   1980
            TabIndex        =   201
            Top             =   780
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DcBranches 
            Height          =   315
            Index           =   3
            Left            =   5280
            TabIndex        =   203
            Top             =   750
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Branch"
            Height          =   195
            Index           =   29
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   810
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Activity"
            Height          =   195
            Index           =   156
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   202
            Top             =   810
            Width           =   735
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   375
            Left            =   9720
            TabIndex        =   192
            Top             =   5250
            Width           =   2055
         End
         Begin VB.Label Label23 
            Caption         =   $"FrmAnalysItems.frx":23FD
            Height          =   375
            Left            =   12120
            TabIndex        =   191
            Top             =   5310
            Width           =   3345
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   375
            Left            =   9720
            TabIndex        =   190
            Top             =   4890
            Width           =   2055
         End
         Begin VB.Label Label21 
            Caption         =   "Total number of unsent invoices"
            Height          =   285
            Left            =   12120
            TabIndex        =   189
            Top             =   5010
            Width           =   2775
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Left            =   9720
            TabIndex        =   188
            Top             =   4650
            Width           =   2055
         End
         Begin VB.Label Label19 
            Caption         =   "Total number of invoices sent"
            Height          =   285
            Left            =   12120
            TabIndex        =   187
            Top             =   4650
            Width           =   2775
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   375
            Left            =   9720
            TabIndex        =   186
            Top             =   4290
            Width           =   2055
         End
         Begin VB.Label Label17 
            Caption         =   "Total invoices"
            Height          =   375
            Left            =   12120
            TabIndex        =   185
            Top             =   4290
            Width           =   2775
         End
         Begin VB.Label Label13 
            Caption         =   "To"
            Height          =   255
            Index           =   4
            Left            =   5400
            TabIndex        =   184
            Top             =   5010
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "From"
            Height          =   255
            Index           =   3
            Left            =   3120
            TabIndex        =   183
            Top             =   5070
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "From"
            Height          =   255
            Index           =   2
            Left            =   7470
            TabIndex        =   182
            Top             =   780
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "TO"
            Height          =   255
            Index           =   0
            Left            =   9630
            TabIndex        =   181
            Top             =   780
            Width           =   375
         End
         Begin VB.Label Label13 
            Caption         =   "Password"
            Height          =   255
            Index           =   1
            Left            =   450
            TabIndex        =   180
            Top             =   210
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   34
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   4560
            Width           =   1740
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   12150
         Index           =   2
         Left            =   -22320
         TabIndex        =   83
         Top             =   45
         Width           =   21675
         Begin VB.OptionButton Optrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ê« Ì— «·„»Ì⁄«  "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnClear 
            Caption         =   "„”Õ"
            Height          =   495
            Index           =   2
            Left            =   2640
            TabIndex        =   93
            Top             =   6360
            Width           =   1335
         End
         Begin VB.Frame Frame7 
            Height          =   7095
            Left            =   5880
            TabIndex        =   91
            Top             =   120
            Width           =   4455
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·”« —Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   615
               Left            =   240
               TabIndex        =   92
               Top             =   6360
               Visible         =   0   'False
               Width           =   3975
            End
            Begin VB.Image Image3 
               Height          =   5715
               Left            =   0
               Picture         =   "FrmAnalysItems.frx":2433
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4395
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·› —Â"
            Height          =   735
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   1920
            Width           =   4455
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   330
               Left            =   2280
               TabIndex        =   87
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   164298755
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   330
               Left            =   120
               TabIndex        =   88
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   164298755
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   15
               Left            =   3690
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   1
               Left            =   1710
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.OptionButton Optrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ê« Ì— "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton Optrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ﬁ»Ê÷« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   360
            Width           =   1335
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   94
            Top             =   6360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   5
            Left            =   1320
            TabIndex        =   95
            Top             =   6360
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ «· ﬁ—Ì—"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin MSDataListLib.DataCombo DcbBranch2 
            Height          =   315
            Left            =   120
            TabIndex        =   96
            Top             =   1440
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCRegionID2 
            Height          =   315
            Left            =   120
            TabIndex        =   101
            Top             =   1080
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCActivity 
            Bindings        =   "FrmAnalysItems.frx":498B
            Height          =   315
            Left            =   120
            TabIndex        =   106
            Top             =   720
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰‘«ÿ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‰ÿﬁ…"
            Height          =   255
            Index           =   17
            Left            =   4485
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   1080
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   26
            Left            =   6525
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   960
            Width           =   2925
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «· ﬁ—Ì— «·„ÿ·Ê»"
            Height          =   255
            Left            =   3840
            TabIndex        =   99
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   20
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   255
            Index           =   18
            Left            =   4485
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   1440
            Width           =   1740
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   12150
         Index           =   0
         Left            =   -22620
         TabIndex        =   57
         Top             =   45
         Width           =   21675
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ÌÊ„Ì… ﬂ«„· „Œ ’— "
            ForeColor       =   &H00FF0000&
            Height          =   345
            Index           =   11
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   207
            Top             =   870
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„Ê·«  "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   10
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   165
            Top             =   1200
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ÌÊ„Ì… «Ã·"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   9
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   1200
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ÌÊ„Ì… ›Ì“«"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   8
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   1200
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ÌÊ„Ì… ‘»ﬂ…"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   7
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ÌÊ„Ì… ‰ﬁœÌ"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ÌÊ„Ì… ﬂ«„· "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ƒ‘—«  «·»Ì⁄  ›’Ì·Ì "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ƒ‘—«  «·»Ì⁄ «Ã„«·Ì"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   360
            Width           =   2175
         End
         Begin VB.Frame Frame8 
            Caption         =   "Õœœ"
            Height          =   615
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   3840
            Width           =   4815
            Begin VB.OptionButton optPos 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬂ·"
               Height          =   255
               Index           =   2
               Left            =   720
               TabIndex        =   114
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optPos 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ··»Ì« "
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   113
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optPos 
               Alignment       =   1  'Right Justify
               Caption         =   "‰ﬁÿÂ ›ﬁÿ"
               Height          =   255
               Index           =   0
               Left            =   3480
               TabIndex        =   112
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· Õ·Ì· »«·›Ê« Ì— "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   360
            Width           =   1455
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   3240
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3630
            TabIndex        =   78
            Top             =   2880
            Width           =   1050
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì «·‘»ﬂ…"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·‘»ﬂ…"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   360
            Width           =   1575
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·› —Â"
            Height          =   1395
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   4560
            Width           =   4455
            Begin MSComCtl2.DTPicker DtpDateFrom2 
               Height          =   330
               Left            =   2280
               TabIndex        =   62
               Top             =   270
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   164364291
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DtpDateTo2 
               Height          =   330
               Left            =   120
               TabIndex        =   63
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   164364291
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker XPDtbTransTimeFrom 
               Height          =   285
               Left            =   2340
               TabIndex        =   122
               Top             =   810
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   164364290
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker XPDtbTransTimeTo 
               Height          =   285
               Left            =   120
               TabIndex        =   124
               Top             =   780
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   164364290
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   270
               Index           =   27
               Left            =   1575
               TabIndex        =   125
               Top             =   810
               Width           =   630
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   270
               Index           =   21
               Left            =   3795
               TabIndex        =   123
               Top             =   840
               Width           =   630
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   12
               Left            =   1710
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   240
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   11
               Left            =   3690
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   240
               Width           =   540
            End
         End
         Begin VB.Frame Frame4 
            Height          =   7095
            Left            =   5880
            TabIndex        =   59
            Top             =   120
            Width           =   4455
            Begin VB.Image Image2 
               Height          =   5715
               Left            =   0
               Picture         =   "FrmAnalysItems.frx":49A0
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4395
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·”« —Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   615
               Index           =   0
               Left            =   240
               TabIndex        =   60
               Top             =   6360
               Visible         =   0   'False
               Width           =   3975
            End
         End
         Begin VB.CommandButton btnClear 
            Caption         =   "„”Õ"
            Height          =   495
            Index           =   1
            Left            =   2640
            TabIndex        =   58
            Top             =   6360
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo DCboStoreName2 
            Height          =   315
            Left            =   120
            TabIndex        =   67
            Top             =   2520
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Top             =   6360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   3
            Left            =   1320
            TabIndex        =   69
            Top             =   6360
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ «· ﬁ—Ì—"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   120
            TabIndex        =   76
            Top             =   2160
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbEmp 
            Height          =   315
            Left            =   120
            TabIndex        =   79
            Top             =   2880
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCRegionID 
            Height          =   315
            Left            =   120
            TabIndex        =   103
            Top             =   1800
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCActivity2 
            Bindings        =   "FrmAnalysItems.frx":6EF8
            Height          =   315
            Left            =   120
            TabIndex        =   108
            Top             =   1440
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰‘«ÿ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‰ÿﬁ…"
            Height          =   255
            Index           =   19
            Left            =   4485
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   1800
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   255
            Index           =   13
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   3240
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·›—⁄ „⁄Ì‰"
            Height          =   255
            Index           =   24
            Left            =   4485
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   2160
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·»«∆⁄"
            Height          =   285
            Index           =   23
            Left            =   4275
            TabIndex        =   74
            Top             =   2940
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   22
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «· ﬁ—Ì— «·„ÿ·Ê»"
            Height          =   255
            Left            =   3840
            TabIndex        =   72
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·„Œ“‰ „⁄Ì‰"
            Height          =   255
            Index           =   16
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   2520
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   14
            Left            =   6525
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   960
            Width           =   2925
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   12150
         Index           =   1
         Left            =   -22920
         TabIndex        =   2
         Top             =   45
         Width           =   21675
         Begin VB.TextBox TxtItemCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   2580
            Width           =   1575
         End
         Begin VB.CommandButton btnClear 
            Caption         =   "„”Õ"
            Height          =   495
            Index           =   0
            Left            =   2640
            TabIndex        =   54
            Top             =   6720
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Height          =   4575
            Left            =   5880
            TabIndex        =   34
            Top             =   120
            Width           =   4455
            Begin VB.Label lblCompanyname 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·”« —Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   615
               Left            =   240
               TabIndex        =   35
               Top             =   3840
               Visible         =   0   'False
               Width           =   3975
            End
            Begin VB.Image Image1 
               Height          =   3675
               Left            =   0
               Picture         =   "FrmAnalysItems.frx":6F0D
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4395
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "„‰ «·› —Â"
            Height          =   735
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   5880
            Width           =   4455
            Begin MSComCtl2.DTPicker DtpDateFrom 
               Height          =   330
               Left            =   2280
               TabIndex        =   30
               Top             =   270
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   168099843
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DtpDateTo 
               Height          =   330
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   168099843
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   4
               Left            =   3690
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   3
               Left            =   1710
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.TextBox ItemDetailedCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   3000
            Width           =   4560
         End
         Begin VB.TextBox ParrtNoCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   3360
            Width           =   4560
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Êﬁ› „Œ“Ê‰"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì „»Ì⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì „‘ —Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   840
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „»Ì⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „‘ —Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   4
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1320
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì „—œÊœ«  „»Ì⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  „—œÊœ«  „‘ —Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „—œÊœ«  „»Ì⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   7
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „—œÊœ«  „‘ —Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   8
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   1320
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’«›Ì «·„»Ì⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   9
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’«›Ì «·„‘ —Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   10
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „»Ì⁄«  Ê„—œÊœ« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   11
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „‘ —Ì«  Ê„—œÊœ« "
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   12
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   1320
            Width           =   2175
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—’Ìœ «›  «ÕÌ «Ã„«·Ì"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   13
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—’Ìœ «›  «ÕÌ  Õ·Ì·Ì"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   14
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.CheckBox ChsERIAL 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»ﬁ« ··”Ì—Ì«·/«·»«—ﬂÊœ"
            Height          =   195
            Left            =   7800
            TabIndex        =   11
            Top             =   5160
            Width           =   2055
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3600
            TabIndex        =   10
            Top             =   4440
            Width           =   1050
         End
         Begin VB.Frame Frame2 
            Height          =   1095
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   4800
            Width           =   4575
            Begin VB.TextBox percent1 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   1920
               TabIndex        =   6
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox percent2 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   240
               TabIndex        =   5
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "‰”»… «·‰„Ê ›Ï «·„»Ì⁄«  "
               Height          =   375
               Left            =   2520
               TabIndex        =   9
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "‰”»… «·«„«‰"
               Height          =   375
               Index           =   0
               Left            =   960
               TabIndex        =   8
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "„ﬁ«—‰Â „⁄ „»Ì⁄« "
               ForeColor       =   &H00C00000&
               Height          =   495
               Index           =   0
               Left            =   1560
               TabIndex        =   7
               Top             =   720
               Width           =   2655
            End
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ﬁ—Ì— „ Œ’’ ··ÿ·»Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   15
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   5040
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo DCboItemsName 
            Height          =   315
            Left            =   75
            TabIndex        =   36
            Top             =   2580
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   75
            TabIndex        =   37
            Top             =   1800
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbColor 
            Height          =   315
            Left            =   75
            TabIndex        =   38
            Top             =   3720
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSize 
            Height          =   315
            Left            =   75
            TabIndex        =   39
            Top             =   4080
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   582
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboGroup1 
            Height          =   315
            Left            =   75
            TabIndex        =   40
            Top             =   2160
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   75
            TabIndex        =   41
            Top             =   4440
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   582
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   55
            Top             =   6720
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   1320
            TabIndex        =   56
            Top             =   6720
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ «· ﬁ—Ì—"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   975
            Left            =   6120
            Top             =   5520
            Width           =   3615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "‘«‘…  ﬁ«—Ì—   Õ·Ì·Ì ··«’‰«›"
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
            Height          =   900
            Index           =   25
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   5520
            Width           =   3495
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   31
            Left            =   6525
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   960
            Width           =   2925
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "’‰› —∆Ì”Ì"
            Height          =   255
            Index           =   30
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   2640
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·„Œ“‰ „⁄Ì‰"
            Height          =   255
            Index           =   8
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   1800
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·ﬂÊœ  Õ·Ì·Ì „⁄Ì‰"
            Height          =   255
            Index           =   0
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   3000
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "··Ê‰ „⁄Ì‰"
            Height          =   255
            Index           =   2
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   3720
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·„ﬁ«” „⁄Ì‰"
            Height          =   255
            Index           =   5
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   4080
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·„Ã„Ê⁄Â „⁄Ì‰…"
            Height          =   255
            Index           =   6
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   2160
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»«—ﬂÊœ Œ«—ÃÌ"
            Height          =   255
            Index           =   7
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   3360
            Width           =   1740
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «· ﬁ—Ì— «·„ÿ·Ê»"
            Height          =   255
            Left            =   3840
            TabIndex        =   44
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   9
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   285
            Index           =   10
            Left            =   4275
            TabIndex        =   42
            Top             =   4500
            Width           =   1365
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   12150
         Index           =   11
         Left            =   22410
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   45
         Width           =   21675
         _cx             =   38232
         _cy             =   21431
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   10860
            TabIndex        =   140
            Top             =   2370
            Width           =   1710
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»ﬁ« · «—ÌŒ"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   480
            Visible         =   0   'False
            Width           =   5265
            Begin VB.OptionButton Rd 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«” Õﬁ«ﬁ"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   240
               Value           =   -1  'True
               Width           =   1245
            End
            Begin VB.OptionButton Rd 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«’œ«— «·›« Ê—…"
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   0
               Left            =   2700
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   240
               Width           =   1605
            End
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”Õ"
            Height          =   555
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Top             =   6000
            Width           =   1560
         End
         Begin VB.TextBox StrCusID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Top             =   7200
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox CurrenrEmployeeIDs 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   7080
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   660
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   -150
            Width           =   17475
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "   ﬁ«—Ì— «⁄„«— «·«’‰«›"
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
               Index           =   22
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   300
               Width           =   3390
            End
         End
         Begin VB.CommandButton CmdSelectEmp 
            Caption         =   " ÕœÌœ>>"
            Height          =   330
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   3150
            Visible         =   0   'False
            Width           =   4320
         End
         Begin VB.CommandButton CmdSelectCus 
            Caption         =   " ÕœÌœ>>"
            Height          =   330
            Left            =   5130
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Top             =   2820
            Width           =   4320
         End
         Begin VB.TextBox txtTotalStill 
            Alignment       =   1  'Right Justify
            Height          =   465
            Left            =   7470
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   8130
            Visible         =   0   'False
            Width           =   2805
         End
         Begin MSComCtl2.DTPicker DTP_Date 
            Height          =   345
            Left            =   105
            TabIndex        =   141
            TabStop         =   0   'False
            Top             =   750
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   211943427
            CurrentDate     =   37140
         End
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   330
            Left            =   4080
            TabIndex        =   142
            Top             =   7080
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   582
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAnalysItems.frx":9465
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Index           =   8
            Left            =   5100
            TabIndex        =   143
            Top             =   1770
            Visible         =   0   'False
            Width           =   7545
            _ExtentX        =   13309
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DCboItemsName2 
            Height          =   315
            Left            =   5130
            TabIndex        =   144
            Top             =   2370
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   582
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   555
            Index           =   0
            Left            =   3960
            TabIndex        =   145
            Top             =   6000
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   979
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…  Õ·Ì·Ì"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAnalysItems.frx":97FF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   555
            Index           =   1
            Left            =   2160
            TabIndex        =   146
            Top             =   6000
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   979
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «Ã„«·Ì"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAnalysItems.frx":9B99
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin XtremeSuiteControls.CheckBox ChekCustomer 
            Height          =   375
            Left            =   13290
            TabIndex        =   147
            Top             =   2400
            Width           =   3075
            _Version        =   786432
            _ExtentX        =   5424
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "’‰›"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CheckAllCustomer 
            Height          =   375
            Left            =   12210
            TabIndex        =   148
            Top             =   2880
            Width           =   4155
            _Version        =   786432
            _ExtentX        =   7329
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "«Œ Ì«— «ﬂÀ— „‰ ’‰›"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox CheckAllEMp 
            Height          =   375
            Left            =   11970
            TabIndex        =   149
            Top             =   4680
            Visible         =   0   'False
            Width           =   4395
            _Version        =   786432
            _ExtentX        =   7752
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "«Œ Ì«— «ﬂÀ— „‰ „‰œÊ»"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker FromDate1 
            Height          =   345
            Left            =   2925
            TabIndex        =   150
            TabStop         =   0   'False
            Top             =   8220
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   211943427
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker DTPickerAccFrom 
            Height          =   345
            Left            =   3120
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   7860
            Visible         =   0   'False
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   211943427
            CurrentDate     =   37140
         End
         Begin VSFlex8Ctl.VSFlexGrid grdAging 
            Height          =   3630
            Left            =   8070
            TabIndex        =   152
            Top             =   5640
            Visible         =   0   'False
            Width           =   8550
            _cx             =   15081
            _cy             =   6403
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   24
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmAnalysItems.frx":9F33
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
         Begin VSFlex8Ctl.VSFlexGrid grdAging2 
            Height          =   1470
            Left            =   1830
            TabIndex        =   153
            Top             =   3900
            Visible         =   0   'False
            Width           =   10050
            _cx             =   17727
            _cy             =   2593
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   21
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmAnalysItems.frx":A2DD
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
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   0
            Left            =   420
            TabIndex        =   154
            Top             =   6030
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontName        =   "Arial"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmAnalysItems.frx":A60F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   166
            Top             =   2280
            Width           =   2235
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ì"
            Height          =   375
            Left            =   4830
            RightToLeft     =   -1  'True
            TabIndex        =   164
            Top             =   8220
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   375
            Left            =   5940
            RightToLeft     =   -1  'True
            TabIndex        =   163
            Top             =   7950
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «’œ«— «·›« Ê—… „‰"
            Height          =   375
            Left            =   5475
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   7500
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   375
            Left            =   1410
            RightToLeft     =   -1  'True
            TabIndex        =   161
            Top             =   7860
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Shape Shape2 
            BorderWidth     =   2
            Height          =   1455
            Left            =   5400
            Top             =   5880
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Â–Â «·‘«‘…  ﬁÊ„ »«ŸÂ«— »Ì«‰«  «⁄„«— «·œÌÊ‰ ÿ»ﬁ« · «—ÌŒ «·ﬁÌ«”"
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
            Height          =   1380
            Index           =   28
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Top             =   6480
            Visible         =   0   'False
            Width           =   6975
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   375
            Left            =   1410
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   7500
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ „‰"
            Height          =   375
            Left            =   5805
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   7860
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "›—⁄ „⁄Ì‰"
            Height          =   375
            Index           =   3
            Left            =   13320
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   1770
            Visible         =   0   'False
            Width           =   3075
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·ﬁÌ«”"
            Height          =   375
            Index           =   3
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   156
            Top             =   720
            Width           =   1350
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«Ã„«·Ì «·„ »ﬁÌ"
            Height          =   375
            Index           =   4
            Left            =   10110
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   7860
            Visible         =   0   'False
            Width           =   1215
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " Õ·Ì·Ì ‰ﬁ«ÿ «·»Ì⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   0
      Left            =   330
      TabIndex        =   0
      Top             =   150
      Width           =   10320
   End
End
Attribute VB_Name = "FrmAnalysItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim ii As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Dim reportid As Integer
Public mIndex As Long
Dim shortFileName As String
Dim excelShortName
Dim excelFileNameFullPath
Dim excelRowNo
Dim POSConnection  As ADODB.Connection

Private Sub ChangeLang()
On Error GoTo ErrTrap

    Label9.Caption = "Activity"
    Label11.Caption = "Activity"
    Label5(0).Caption = "Items Analysis Report"
    lbl(25).Caption = Label5(0).Caption
    Label2.Caption = "Select Reports"
    lblCompanyname.Caption = "AL SATTARYAH"
    lbl(2).Caption = "Color"
    lbl(5).Caption = "Size"
    ChsERIAL.Caption = "Serials/barcode"
    lbl(19).Caption = "Region"
    lbl(17).Caption = "Region"
    lbl(4).Caption = "From"
    lbl(3).Caption = "To"
    Frame1.Caption = "Period"
    lbl(8).Caption = "Store"
    lbl(6).Caption = "Group"
    lbl(30).Caption = "Item"
    lbl(0).Caption = "Code"
    lbl(7).Caption = "BarCode"
    opt(0).RightToLeft = False
    opt(1).RightToLeft = False
    opt(2).RightToLeft = False
    opt(3).RightToLeft = False
    opt(4).RightToLeft = False
    opt(5).RightToLeft = False
    opt(6).RightToLeft = False
    opt(7).RightToLeft = False
    opt(8).RightToLeft = False
    opt(9).RightToLeft = False
    opt(10).RightToLeft = False
    opt(11).RightToLeft = False
    opt(12).RightToLeft = False
    opt(13).RightToLeft = False
    opt(14).RightToLeft = False
    'opt(15).RightToLeft = False
    btnClear(0).Caption = "Clear"
    btnClear(1).Caption = "Clear"
    Cmd(0).Caption = "Show"
    Cmd(2).Caption = "Exit"
    opt(0).Caption = "Inventory Stock"
    opt(1).Caption = "Total Sales"
    opt(2).Caption = "Total Purchases"
    opt(3).Caption = "Analytical Sales"
    opt(4).Caption = "Analytical Purchases"
    opt(13).Caption = "Total Op Balance "
    opt(14).Caption = "Analytical Op Balance "
    opt(5).Caption = "Total Sales Returns "
    opt(6).Caption = "Total Purchases Returns "
    opt(7).Caption = "Anal. Sales Returns "
    opt(8).Caption = "Anal. Purchases Returns "
    opt(9).Caption = "Net Sales "
    opt(10).Caption = "Net Purchases "
    opt(11).Caption = "Analy.Sales and Returns "
    opt(12).Caption = "Analy.Purchases and Returns "
ErrTrap:
End Sub

Private Sub btnClear_Click(Index As Integer)
    clear_all Me
    DtpDateFrom.value = ""
    DtpDateTo.value = ""
    DtpDateFrom2.value = ""
    DtpDateTo2.value = ""
    DTPicker2.value = ""
    DTPicker1.value = ""
    optNetWork(0).value = True
End Sub

Private Sub btnNew_Click(Index As Integer)
   On Error GoTo eh
        CommonDialog1.DialogTitle = "Select Upload list"
        CommonDialog1.CancelError = True
        CommonDialog1.filter = "xls Files (*.xls)|*.xls"
        CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNLongNames Or cdlOFNExplorer
        CommonDialog1.ShowOpen
        Dim vFiles() As String
       
        
        'grdFiles.Visible = True
    
        '  grdFiles.rows = 1
        Dim Row As Integer
        Row = 1
        Dim i As Integer
        vFiles = Split(CommonDialog1.FileName, CHR(0))
        If UBound(vFiles) = 0 Then
               ' OpenExcelFileAndFormatDates CommonDialog1.FileName
                'SaveExcelFile2 CommonDialog1.FileName
            FetchExcelDataWithZatcaStatus CommonDialog1.FileName
           
        End If
eh:
    
End Sub

Public Sub loadgridExcel(ByVal Sqlstmt As String, ByRef tGrd As Control, Optional ResetRows As Boolean = True, Optional InsertRow As Boolean = False, Optional mReCreateColumns As Boolean = False, Optional ByVal Conn As ADODB.Connection = Nothing)
    Dim tRs As New ADODB.Recordset

    If Conn Is Nothing Then
        tRs.Open Sqlstmt, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Else
        tRs.Open Sqlstmt, Conn, adOpenKeyset, adLockReadOnly
    End If

    ' Add your logic to populate the grid with tRs here
    ' Example:
    If Not tRs.EOF Then
        tGrd.rows = 1
        tGrd.Cols = tRs.Fields.count
        Do While Not tRs.EOF
            tGrd.AddItem tRs.GetString(adClipString, 1, vbTab, vbCrLf)
            tRs.MoveNext
        Loop
    End If

    tRs.Close
    Set tRs = Nothing
End Sub


Private Sub OpenExcelFileAndFormatDates(ByVal FileName As String)
    If FileName = "" Then
        MsgBox "«Œ — „·› «Ê·«"
        Exit Sub
    End If

    Dim moConn As New ADODB.Connection
    Dim mrs As ADODB.Recordset
    Dim sfo As New FileSystemObject
    
    shortFileName = sfo.GetFileName(FileName)

    ' › Õ «·« ’«· »„·› Excel
    moConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source='" & FileName & "'; Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
    Set mrs = moConn.OpenSchema(adSchemaTables)

    '  ÕœÌœ Ê—ﬁ… «·⁄„·
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim cell As Object
    Dim lastrow As Long

    ' ≈‰‘«¡ ﬂ«∆‰ Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(FileName)
    Set xlSheet = xlBook.Sheets("SOMATCO TO ZATCA REPORT") ' «” »œ· "Sheet1" »«”„ Ê—ﬁ… «·⁄„· «·Œ«’… »ﬂ

    ' «·⁄ÀÊ— ⁄·Ï ¬Œ— ’›
    lastrow = xlSheet.cells(xlSheet.rows.count, "A").End(-4162).Row ' -4162 ÂÊ xlUp ›Ì VB6

    '  ÕÊÌ·  ‰”Ìﬁ «· Ê«—ÌŒ
    For Each cell In xlSheet.Range("C1:C" & lastrow) ' «” »œ· "A1:A" »«·⁄„Êœ «·–Ì ÌÕ ÊÌ ⁄·Ï «· Ê«—ÌŒ
        If IsDate(cell.value) Then
            cell.value = Format(cell.value, "yyyy-mm-dd")
            cell.value = Format(cell.value, "yyyy-dd-mm")
        End If
    Next cell

    ' Õ›Ÿ Ê≈€·«ﬁ „·› Excel
    xlBook.save
    xlBook.Close False
    xlApp.Quit

    '  ‰ŸÌ› «·ﬂ«∆‰« 
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    ' «” ﬂ„«· «·⁄„·Ì«  «·√Œ—Ï
    ' ...

    MsgBox " „  ÕÊÌ·  ‰”Ìﬁ «· Ê«—ÌŒ »‰Ã«Õ!"
End Sub


Private Sub FetchExcelDataWithZatcaStatus(FileName)
   On Error GoTo ErrorHandler
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim FilePath As String
    Dim moConn As New ADODB.Connection
    Dim mrs As ADODB.Recordset
    Dim RsData As New ADODB.Recordset
    Dim sfo As New FileSystemObject
    Dim shortFileName As String
    Dim tblname As String
    Dim zatcaStatusColExists As Boolean
    Dim Col As Long
Dim s As String
    '  ÕœÌœ „”«— „·› Excel
    FilePath = FileName ' «” »œ· «·„”«— »„”«— „·› Excel «·Œ«’ »ﬂ
excelFileNameFullPath = FileName
    ' ≈‰‘«¡ ﬂ«∆‰ Excel ÃœÌœ »«” Œœ«„ Late Binding
    Set xlApp = CreateObject("Excel.Application")

    ' › Õ „·› Excel
    Set xlBook = xlApp.Workbooks.Open(FilePath)

    '  ÕœÌœ «·‘Ì  «·√Ê· »€÷ «·‰Ÿ— ⁄‰ «”„Â
    Set xlSheet = xlBook.Sheets(1)

    ' «·»ÕÀ ⁄‰ ⁄„Êœ zatcaStatus
    zatcaStatusColExists = False
    For Col = 1 To xlSheet.UsedRange.Columns.count
        If xlSheet.cells(1, Col).value = "zatcaStatus" Then
            zatcaStatusColExists = True
            Exit For
        End If
    Next Col

    ' ≈€·«ﬁ „·› Excel »⁄œ «· Õﬁﬁ „‰ ÊÃÊœ «·⁄„Êœ
    xlBook.Close False
    xlApp.Quit
Dim rsDummy
    '  ‰ŸÌ› «·ﬂ«∆‰« 
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    ' › Õ « ’«· »„·› Excel »«” Œœ«„ ADODB
    moConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source='" & FilePath & "'; Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"


'moConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'            "Data Source='" & FilePath & "';" & _
'            "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
    ' «” œ⁄«¡ „⁄·Ê„«  «·Ãœ«Ê·
    Set mrs = moConn.OpenSchema(adSchemaTables)

    ' ÷»ÿ «·»Ì«‰«  ›Ì «·ÃœÊ·
    If Not mrs.EOF Then
        tblname = mrs.Fields("table_name").value
        RsData.CursorLocation = adUseClient
        If zatcaStatusColExists Then
            'RsData.Open "SELECT * FROM [" & tblname & "] WHERE zatcaStatus IS NULL OR zatcaStatus = ''", moConn, adOpenKeyset, adLockReadOnly
            's = "SELECT * FROM [" & tblname & "]   WHERE ISNULL(zatcaStatus, '') = ''"
            s = "SELECT * FROM [" & tblname & "] WHERE  zatcaStatus = 0 "
            

        Else
            'RsData.Open "SELECT * FROM [" & tblname & "]", moConn, adOpenKeyset, adLockReadOnly
            s = "SELECT * FROM [" & tblname & "] "
        End If
        
            
    

    ' ⁄—÷ «·»Ì«‰«  ›Ì «·ÃœÊ· („À«·° «” »œ· Â–« »«·„⁄«·Ã… «·›⁄·Ì…)
'    While Not RsData.EOF
'        Debug.Print RsData.Fields("invoiceid").value
'        RsData.MoveNext
'    Wend



' Dim tRs As New ADODB.Recordset
'
'
'        tRs.Open s, moConn, adOpenKeyset, adLockReadOnly
'
'
'    ' Add your logic to populate the grid with tRs here
'    ' Example:
'    If Not tRs.EOF Then
'        grd(0).rows = 1
'        grd(0).Cols = tRs.Fields.count
'        Do While Not tRs.EOF
'            grd(0).AddItem tRs.GetString(adClipString, 1, vbTab, vbCrLf)
'            tRs.MoveNext
'        Loop
'    End If
'
'    tRs.Close
'    Set tRs = Nothing


        
        
        
        
        
        Set rsDummy = New ADODB.Recordset
    

    rsDummy.Open s, moConn, adOpenKeyset, adLockReadOnly
    grd(0).rows = 1
    grd(0).rows = grd(0).rows + 1
    With grd(0)
      .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
       ' .AutoResize = True
        .AutoSizeMode = flexAutoSizeColWidth
        'Set .WallPaper = Grdback.MoneyWallpaper
  .ColComboList(.ColIndex("viewFile")) = "..."
  .ColComboList(.ColIndex("ErrorMessage")) = "..."
  .ColComboList(.ColIndex("warrningmessage")) = "..."
  
  .ColComboList(.ColIndex("viewQRCode")) = "..."
  .ColComboList(.ColIndex("ViewError")) = "..."
    .ColComboList(.ColIndex("View")) = "..."
  
       ' .AutoSize 0,  .Cols - 1, False
 
     
     End With
     
     
         With grd(2)
      .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
       ' .AutoResize = True
        .AutoSizeMode = flexAutoSizeColWidth
        'Set .WallPaper = Grdback.MoneyWallpaper
  .ColComboList(.ColIndex("viewFile")) = "..."
  .ColComboList(.ColIndex("ErrorMessage")) = "..."
  .ColComboList(.ColIndex("viewQRCode")) = "..."
  .ColComboList(.ColIndex("warrningmessage")) = "..."
      .ColComboList(.ColIndex("View")) = "..."
       ' .AutoSize 0,  .Cols - 1, False
 
     

     End With
     
     
    Dim i As Long
    Dim mTotalNet As Double
    Dim mTotalDiscountNet As Double
    Dim mTransaction_NetValue As Double
    Dim ReturnSerial As String
Dim SalesInvoiceDate As String
    i = grd(0).rows - 1

Dim rsBranch As New ADODB.Recordset
Dim OtherInformation As New ClsGLOther

Dim StrDate As String

Dim yyyy As String, mm As String, dd As String
Dim mMasterGui As String
mMasterGui = GenerateGUID
Dim rsDummy2 As New ADODB.Recordset

Dim hasNewNO As Boolean
Dim hasManualInvoiceNo As Boolean
Dim hasIqarName As Boolean
Dim hasComResid As Boolean

hasNewNO = FieldExists(rsDummy, "NewNO")
hasManualInvoiceNo = FieldExists(rsDummy, "ManualInvoiceNo")
hasIqarName = FieldExists(rsDummy, "IqarName")
hasComResid = FieldExists(rsDummy, "ComResid")

    Do While Not rsDummy.EOF
        s = "Select InvoiceID from tblEInvoice Where InvoiceID =  N'" & Trim(rsDummy("InvoiceID")) & "'"
        Set rsDummy2 = New ADODB.Recordset
        rsDummy2.Open s, Cn, adOpenForwardOnly, adLockReadOnly
        If Not rsDummy2.EOF Then
            GoTo NextRow
        End If
        
 
         grd(0).TextMatrix(i, grd(0).ColIndex("GroupUniqueFileMaster")) = mMasterGui
        Dim e As New ClsGLOther
         
        e.Invoicetype = IIf(IsNull(rsDummy("DefaultInvoicetype").value), 0, rsDummy("DefaultInvoicetype").value)
          
        e.ErrorMessageS = ""
          

   
   
        grd(0).TextMatrix(i, grd(0).ColIndex("Ser")) = i
        grd(0).TextMatrix(i, grd(0).ColIndex("docType")) = 10
        grd(0).TextMatrix(i, grd(0).ColIndex("ErrorMessage")) = "" ' IIf(IsNull(rsDummy("ErrorMessages").value), "", rsDummy("ErrorMessages").value)
        
        grd(0).TextMatrix(i, grd(0).ColIndex("Id700")) = "" 'IIf(IsNull(rsDummy("Id700").value), "", rsDummy("Id700").value)
        
        
        grd(0).TextMatrix(i, grd(0).ColIndex("chkTaxExempt")) = e.chkTaxExempt
        
        grd(0).TextMatrix(i, grd(0).ColIndex("Export")) = IIf(IsNull(rsDummy("Export").value), "", rsDummy("Export").value)
        
        
        If FieldExists(rsDummy, "BranchName") Then
            grd(0).TextMatrix(i, grd(0).ColIndex("BranchName")) = _
            IIf(IsNull(rsDummy("BranchName").value), "", rsDummy("BranchName").value)
        End If

        
        'grd(0).TextMatrix(i, grd(0).ColIndex("BranchName")) = IIf(IsNull(rsDummy("BranchName").value), "", rsDummy("BranchName").value)
        
        s = "Select branch_id from TblBranchesData where branch_name Like N'" & Trim(grd(0).TextMatrix(i, grd(0).ColIndex("BranchName"))) & "' Or branch_namee Like N'" & Trim(grd(0).TextMatrix(i, grd(0).ColIndex("BranchName"))) & "'"
        Set rsBranch = New ADODB.Recordset
        rsBranch.Open s, Cn, adOpenKeyset, adLockReadOnly
        If Not rsBranch.EOF Then
            grd(0).TextMatrix(i, grd(0).ColIndex("branch_id")) = rsBranch!branch_id & ""
        End If
        rsBranch.Close
       
        If hasNewNO Then
    
            grd(0).TextMatrix(i, grd(0).ColIndex("NewNO")) = IIf(IsNull(rsDummy("NewNO").value), "", rsDummy("NewNO").value)
        End If

        If hasManualInvoiceNo Then
            
            grd(0).TextMatrix(i, grd(0).ColIndex("ManualInvoiceNo")) = IIf(IsNull(rsDummy("ManualInvoiceNo").value), "", rsDummy("ManualInvoiceNo").value)
        End If

        If hasIqarName Then
            
            grd(0).TextMatrix(i, grd(0).ColIndex("IqarName")) = IIf(IsNull(rsDummy("IqarName").value), "", rsDummy("IqarName").value)
        End If
        
        If hasComResid Then
            grd(0).TextMatrix(i, grd(0).ColIndex("ComResid")) = IIf(IsNull(rsDummy("ComResid").value), "", rsDummy("ComResid").value)
            
        End If
        
        
        grd(0).TextMatrix(i, grd(0).ColIndex("Transaction_ID")) = IIf(IsNull(rsDummy("InvoiceID").value), "", rsDummy("InvoiceID").value)
        grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceID")) = IIf(IsNull(rsDummy("InvoiceID").value), "", rsDummy("InvoiceID").value)
       
        
        grd(0).TextMatrix(i, grd(0).ColIndex("id")) = IIf(IsNull(rsDummy("InvoiceID").value), "", rsDummy("InvoiceID").value)
        grd(0).TextMatrix(i, grd(0).ColIndex("DefaultInvoicetype")) = e.Invoicetype
        
        grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = IIf(IsNull(rsDummy("IssueDate").value), "", rsDummy("IssueDate").value)
        If IsNull(rsDummy("IssueDate").value) Or rsDummy("IssueDate").value = "" Then
            grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = ""
        Else
            grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = Format$(rsDummy("IssueDate").value, "yyyy-mm-dd")
        End If

StrDate = rsDummy("IssueDate").value ' ·«“„ ÌﬂÊ‰ ‰’ »«·‘ﬂ· yyyy-mm-dd
'
'If IsNull(strDate) Or strDate = "" Then
'    grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = ""
'Else
'    yyyy = left(strDate, 4)
'    mm = mId(strDate, 6, 2)
'    dd = right(strDate, 2)
'    grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = yyyy & "-" & mm & "-" & dd
'End If
'
'Dim arrDate() As String
'
'If Not IsNull(strDate) And Trim(strDate) <> "" Then
'    arrDate = Split(strDate, "-")
'    If UBound(arrDate) = 2 Then
'        ' ????? - ????? - ?????
'        grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = arrDate(0) & "-" & arrDate(1) & "-" & arrDate(2)
'    Else
'        ' ?? ?? ???? ????
'        grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = strDate
'    End If
'Else
'    grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = ""
'End If

Dim arrDate() As String
'Dim strDate As String

StrDate = rsDummy("IssueDate").value

If Not IsNull(StrDate) And Trim(StrDate) <> "" Then
    If InStr(StrDate, "/") > 0 Then
        ' ’Ì€…: dd/mm/yyyy
        arrDate = Split(StrDate, "/")
        If UBound(arrDate) = 2 Then
            ' Â‰«: «·”‰… - «·‘Â— - «·ÌÊ„
            grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = _
                Format$(val(arrDate(2)), "0000") & "-" & _
                Format$(val(arrDate(1)), "00") & "-" & _
                Format$(val(arrDate(0)), "00")
        Else
            grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = StrDate
        End If
    ElseIf InStr(StrDate, "-") > 0 Then
        ' ’Ì€…: yyyy-mm-dd
        arrDate = Split(StrDate, "-")
        If UBound(arrDate) = 2 Then
            ' Â‰«: «·”‰… - «·‘Â— - «·ÌÊ„ »«·›⁄·
            grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = _
                Format$(val(arrDate(0)), "0000") & "-" & _
                Format$(val(arrDate(1)), "00") & "-" & _
                Format$(val(arrDate(2)), "00")
        Else
            grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = StrDate
        End If
    Else
        grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = StrDate
    End If
Else
    grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = ""
End If

grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = ParseAnyDate(rsDummy("IssueDate").value)

        grd(0).TextMatrix(i, grd(0).ColIndex("IssueTim")) = IIf(IsNull(rsDummy("IssueTim").value), "", rsDummy("IssueTim").value)
        
        If e.Invoicetype = 0 Then
            grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceTypeCodename")) = "0100000"
            grd(0).TextMatrix(i, grd(0).ColIndex("PaymentMeansCode")) = 30
            grd(0).TextMatrix(i, grd(0).ColIndex("paymentnote")) = "Payment by Credit"
        
        Else
            grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceTypeCodename")) = "0200000"
            grd(0).TextMatrix(i, grd(0).ColIndex("PaymentMeansCode")) = 10
            grd(0).TextMatrix(i, grd(0).ColIndex("paymentnote")) = "Payment by Cash"
        End If
        grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceTypeCodeID")) = 388
        
        grd(0).TextMatrix(i, grd(0).ColIndex("Qty")) = Trim(rsDummy!Qty & "")
        grd(0).TextMatrix(i, grd(0).ColIndex("Price")) = Trim(rsDummy!Price & "")
        
        grd(0).TextMatrix(i, grd(0).ColIndex("ItemName")) = Trim(rsDummy!ItemName & "")
        
    
       
        
     
        
        grd(0).TextMatrix(i, grd(0).ColIndex("DocumentCurrencyCode")) = "SAR"
        
        grd(0).TextMatrix(i, grd(0).ColIndex("TaxCurrencyCode")) = "SAR"
        grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceDocumentReferenceID")) = ""
        grd(0).TextMatrix(i, grd(0).ColIndex("AdditionalDocumentReferenceICVUUID")) = ""
        grd(0).TextMatrix(i, grd(0).ColIndex("ActualDeliveryDate")) = IIf(IsNull(rsDummy("IssueDate").value), "", rsDummy("IssueDate").value)
        grd(0).TextMatrix(i, grd(0).ColIndex("LatestDeliveryDate")) = IIf(IsNull(rsDummy("IssueDate").value), "", rsDummy("IssueDate").value)
        
        grd(0).TextMatrix(i, grd(0).ColIndex("PayeeFinancialAccount")) = ""
        
        grd(0).TextMatrix(i, grd(0).ColIndex("Identificationid")) = IIf(IsNull(rsDummy("Identificationid").value), "", rsDummy("Identificationid").value)
        
        
        grd(0).TextMatrix(i, grd(0).ColIndex("schemeID")) = "CRN"
        grd(0).TextMatrix(i, grd(0).ColIndex("StreetName")) = Trim(rsDummy!StreetName & "")
        grd(0).TextMatrix(i, grd(0).ColIndex("AdditionalStreetName")) = ""
        grd(0).TextMatrix(i, grd(0).ColIndex("BuildingNumber")) = Trim(rsDummy!BuildingNumber & "")
        grd(0).TextMatrix(i, grd(0).ColIndex("PlotIdentification")) = ""
        
        grd(0).TextMatrix(i, grd(0).ColIndex("CityName")) = Trim(rsDummy!CityName & "")
        grd(0).TextMatrix(i, grd(0).ColIndex("PostalZone")) = Trim(rsDummy!PostalZone & "")
        grd(0).TextMatrix(i, grd(0).ColIndex("CountrySubentity")) = 1
        grd(0).TextMatrix(i, grd(0).ColIndex("CitySubdivisionName")) = Trim(rsDummy!CitySubdivisionName & "")
        grd(0).TextMatrix(i, grd(0).ColIndex("IdentificationCode")) = "SA"
        
        
        grd(0).TextMatrix(i, grd(0).ColIndex("RegistrationName")) = Trim(rsDummy!RegistrationName & "")
        grd(0).TextMatrix(i, grd(0).ColIndex("CompanyID")) = Trim(rsDummy!CompanyID & "")
        '«·Œ’Ê„« 
        grd(0).TextMatrix(i, grd(0).ColIndex("allowancechargeAmount")) = 0
        grd(0).TextMatrix(i, grd(0).ColIndex("AllowanceChargeReason")) = ""
        
        grd(0).TextMatrix(i, grd(0).ColIndex("VATValue")) = Trim(rsDummy!VATValue & "")
       ' grd(0).TextMatrix(i, grd(0).ColIndex("VATValue")) =  Trim(rsDummy!VATValue & "")
        
        '
        If val(rsDummy!VATValue & "") = 0 Then
            grd(0).TextMatrix(i, grd(0).ColIndex("TaxCategoryID")) = "Z"
            grd(0).TextMatrix(i, grd(0).ColIndex("TaxCategoryPercent")) = 0
            grd(0).TextMatrix(i, grd(0).ColIndex("ComResid")) = 0
            e.chkTaxExempt = 1
            grd(0).TextMatrix(i, grd(0).ColIndex("chkTaxExempt")) = e.chkTaxExempt
        Else
            grd(0).TextMatrix(i, grd(0).ColIndex("TaxCategoryID")) = "S"
            grd(0).TextMatrix(i, grd(0).ColIndex("TaxCategoryPercent")) = 15
            grd(0).TextMatrix(i, grd(0).ColIndex("ComResid")) = 1
            e.chkTaxExempt = 0
            grd(0).TextMatrix(i, grd(0).ColIndex("chkTaxExempt")) = e.chkTaxExempt
        End If
        
        
        
        
        grd(0).TextMatrix(i, grd(0).ColIndex("PayableAmount")) = Trim(rsDummy!PayableAmount & "")
        grd(0).TextMatrix(i, grd(0).ColIndex("PrepaidAmount")) = 0
        
        '
        '   grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceHash")) = e.InvoiceHash
        '   grd(0).TextMatrix(i, grd(0).ColIndex("SingedXML")) = e.SingedXML
        '  grd(0).TextMatrix(i, grd(0).ColIndex("EncodedInvoice")) = e.EncodedInvoice
        '   grd(0).TextMatrix(i, grd(0).ColIndex("UUID")) = e.UUID
        '   grd(0).TextMatrix(i, grd(0).ColIndex("QRCode")) = e.QRCode
        '  grd(0).TextMatrix(i, grd(0).ColIndex("PIH")) = e.PIH
        '   grd(0).TextMatrix(i, grd(0).ColIndex("SingedXMLFileName")) = e.SingedXMLFileName
        'grd(0).TextMatrix(i, grd(0).ColIndex("QrCodeDataPath")) = e.QrCodeDataPath
        
        
        
        ' e.generateInvoice
     
       
        i = i + 1
        grd(0).rows = grd(0).rows + 1
NextRow:
        rsDummy.MoveNext
    Loop

        
        
        
        
        
        
        
        
        
        
        
        
        
        
       ' loadgrid s, grd(0), True, False, False, moConn
                           
      Cn.Execute " delete                        tmptblEInvoice"
                          
        
                           
        

        
        
        
       s = "Select * from tmptblEInvoice where id =0 "
        saveGrid s, grd(0), "InvoiceId", "serial"

                           
                           
        s = "Select * from tblEInvoice2 where id =0 "
        saveGrid s, grd(0), "InvoiceId", "serial"
        
      
        Dim sql As String

sql = ""
sql = sql & "IF OBJECT_ID('tempdb..#TmpGroups') IS NOT NULL DROP TABLE #TmpGroups;" & vbCrLf

sql = sql & "SELECT InvoiceID, NEWID() AS NewGroupCode " & vbCrLf
sql = sql & "INTO #TmpGroups " & vbCrLf
sql = sql & "FROM tblEInvoice2 " & vbCrLf
sql = sql & "WHERE GroupUniqueFileMaster = '" & Trim(mMasterGui) & "' " & vbCrLf
sql = sql & "GROUP BY InvoiceID;" & vbCrLf

sql = sql & "UPDATE t " & vbCrLf
sql = sql & "SET t.GroupUniqueCode = g.NewGroupCode " & vbCrLf
sql = sql & "FROM tblEInvoice2 t " & vbCrLf
sql = sql & "INNER JOIN #TmpGroups g ON t.InvoiceID = g.InvoiceID " & vbCrLf
sql = sql & "WHERE t.GroupUniqueFileMaster = '" & Trim(mMasterGui) & "';" & vbCrLf

' (???????) ??? ?????? ?????? ??????? ??? ???????
sql = sql & "DROP TABLE #TmpGroups;" & vbCrLf

' ??? ?????
Cn.Execute sql


  sql = ""
sql = sql & "IF OBJECT_ID('tempdb..#TmpGroups') IS NOT NULL DROP TABLE #TmpGroups;" & vbCrLf

sql = sql & "SELECT InvoiceID, NEWID() AS NewGroupCode " & vbCrLf
sql = sql & "INTO #TmpGroups " & vbCrLf
sql = sql & "FROM tmptblEInvoice " & vbCrLf
sql = sql & "WHERE GroupUniqueFileMaster = '" & Trim(mMasterGui) & "' " & vbCrLf
sql = sql & "GROUP BY InvoiceID;" & vbCrLf

sql = sql & "UPDATE t " & vbCrLf
sql = sql & "SET t.GroupUniqueCode = g.NewGroupCode " & vbCrLf
sql = sql & "FROM tmptblEInvoice t " & vbCrLf
sql = sql & "INNER JOIN #TmpGroups g ON t.InvoiceID = g.InvoiceID " & vbCrLf
sql = sql & "WHERE t.GroupUniqueFileMaster = '" & Trim(mMasterGui) & "';" & vbCrLf

' (???????) ??? ?????? ?????? ??????? ??? ???????
sql = sql & "DROP TABLE #TmpGroups;" & vbCrLf

' ??? ?????
Cn.Execute sql
s = ""
s = s & "UPDATE tmptblEInvoice " & vbCrLf
s = s & "SET chkTaxExempt = CASE " & vbCrLf
s = s & "    WHEN VATValue IS NULL OR VATValue = 0 THEN 1 " & vbCrLf
s = s & "    ELSE 0 " & vbCrLf
s = s & "END " & vbCrLf
Cn.Execute s
      
    
s = ""
s = s & "SELECT " & vbCrLf
s = s & "    InvoiceID,GroupUniqueCode,GroupUniqueFileMaster,NewNO,ComResid,IqarName,ManualInvoiceNo," & vbCrLf
s = s & "    InvoiceID as Transaction_ID,chkTaxExempt," & vbCrLf
s = s & "    Identificationid as Identificationid," & vbCrLf

s = s & "    DefaultInvoiceType," & vbCrLf
s = s & "    ExcelRow," & vbCrLf
s = s & "    ExcelFile," & vbCrLf
s = s & "    IssueDate," & vbCrLf
s = s & "    IssueTim," & vbCrLf
s = s & "    DocumentCurrencyCode,PaymentMeansCode,paymentnote,TaxCategoryID,TaxCategoryPercent,DocumentCurrencyCode," & vbCrLf
s = s & "    TaxCurrencyCode," & vbCrLf
s = s & "    StreetName," & vbCrLf
s = s & "    BuildingNumber," & vbCrLf
s = s & "    CityName," & vbCrLf
s = s & "    PostalZone," & vbCrLf
s = s & "    CitySubdivisionName," & vbCrLf
s = s & "    RegistrationName," & vbCrLf
s = s & "    CompanyID," & vbCrLf
s = s & "    CoCRCode," & vbCrLf
s = s & "    SUM(PayableAmount) AS PayableAmount," & vbCrLf
s = s & "    SUM(VatValue) AS VatValue," & vbCrLf
's = s & "    round(SUM(VatValue) /(SUM(Price*qty)   ) * 100,1)  AS TaxCategoryPercent ," & vbCrLf

s = s & "    TaxCategoryPercent = case when SUM(Price*qty)     > 0 then"
s = s & "        round(SUM(VatValue) /(SUM(Price*qty) ) * 100,1) " & vbCrLf
s = s & "        else 0 end ," & vbCrLf
'15' AS TaxCategoryPercent
s = s & "    Id700," & vbCrLf
s = s & "    QrCodeData," & vbCrLf
s = s & "    QrCodeDataPath," & vbCrLf
s = s & "    zatcaStatus," & vbCrLf
s = s & "    InvoiceTypeCodeID," & vbCrLf
s = s & "    InvoiceTypeCodename," & vbCrLf
s = s & "    AdditionalDocumentReferencePIH," & vbCrLf
s = s & "    InvoiceDocumentReferenceID," & vbCrLf
s = s & "    AdditionalDocumentReferenceICVUUID," & vbCrLf
s = s & "    ActualDeliveryDate," & vbCrLf
s = s & "    LatestDeliveryDate," & vbCrLf
s = s & "    RecTime, " & vbCrLf
s = s & "    Export, " & vbCrLf
s = s & "    tmptblEInvoice.branch_id, " & vbCrLf
s = s & "    branch_name, " & vbCrLf
s = s & "    branchname " & vbCrLf

s = s & " FROM tmptblEInvoice " & vbCrLf
s = s & " GROUP BY " & vbCrLf
s = s & "    InvoiceID," & vbCrLf
s = s & "    DefaultInvoiceType," & vbCrLf
s = s & "    ExcelRow," & vbCrLf
s = s & "    ExcelFile," & vbCrLf
s = s & "    IssueDate," & vbCrLf
s = s & "    IssueTim," & vbCrLf
s = s & "    DocumentCurrencyCode," & vbCrLf
s = s & "    TaxCurrencyCode," & vbCrLf
s = s & "    StreetName," & vbCrLf
s = s & "    BuildingNumber," & vbCrLf
s = s & "    CityName," & vbCrLf
s = s & "    PostalZone," & vbCrLf
s = s & "    CitySubdivisionName," & vbCrLf
s = s & "    RegistrationName," & vbCrLf
s = s & "    CompanyID," & vbCrLf
s = s & "    CoCRCode," & vbCrLf
s = s & "    Id700," & vbCrLf
s = s & "    QrCodeData," & vbCrLf
s = s & "    QrCodeDataPath," & vbCrLf
s = s & "    zatcaStatus," & vbCrLf
s = s & "    InvoiceTypeCodeID," & vbCrLf
s = s & "    InvoiceTypeCodename," & vbCrLf
s = s & "    AdditionalDocumentReferencePIH," & vbCrLf
s = s & "    InvoiceDocumentReferenceID," & vbCrLf
s = s & "    AdditionalDocumentReferenceICVUUID," & vbCrLf
s = s & "    ActualDeliveryDate," & vbCrLf
s = s & "    LatestDeliveryDate,Identificationid,DocumentCurrencyCode,PaymentMeansCode,paymentnote,TaxCategoryID,TaxCategoryPercent,DocumentCurrencyCode,   "
   
s = s & "    RecTime,"
s = s & "    Export, "
s = s & "    tmptblEInvoice.branch_id, " & vbCrLf
s = s & "    branch_name, " & vbCrLf
s = s & "    branchname,GroupUniqueCode,GroupUniqueFileMaster,NewNO,ComResid,chkTaxExempt,IqarName,ManualInvoiceNo " & vbCrLf
        
        loadgrid s, grd(0), True, False
        
        
         s = "Select * from tblEInvoice where id =0 "
        saveGrid s, grd(0), "InvoiceId", "serial"
        
        s = "Select * from tblEInvoice where InvoiceId In (Select InvoiceId from tmptblEInvoice) "
        Set rsDummy = New ADODB.Recordset
        
        rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
        Do While Not rsDummy.EOF
            
        SaveQRCode "tblEInvoice", "ID", val(rsDummy!ID & ""), Trim(rsDummy!invoiceID & ""), rsDummy!IssueDate & "", _
        val(rsDummy!PayableAmount & ""), Picture1, 0, val(rsDummy!VATValue & ""), val(rsDummy!PayableAmount & "")
            
            rsDummy.MoveNext
        Loop
                         
End If
        
    ' ≈€·«ﬁ «·« ’«·
   ' RsData.Close
    moConn.Close

    '  ‰ŸÌ› «·ﬂ«∆‰« 
    Set RsData = Nothing
    Set mrs = Nothing
    Set moConn = Nothing

    MsgBox "Data fetched successfully!"
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set mrs = Nothing
    If Not moConn Is Nothing Then moConn.Close
    Set moConn = Nothing
End Sub

Public Function FieldExists(ByVal rs As ADODB.Recordset, FieldName As String) As Boolean
    Dim fld As ADODB.Field
    FieldExists = False
    For Each fld In rs.Fields
        If LCase(fld.Name) = LCase(FieldName) Then
            FieldExists = True
            Exit Function
        End If
    Next
End Function

Function ParseAnyDate(StrDate As String) As String
    Dim dayStr As String, MonthStr As String, YearStr As String
    Dim parts() As String
    Dim i As Integer

    StrDate = Trim(StrDate)
    If StrDate = "" Or IsNull(StrDate) Then
        ParseAnyDate = ""
        Exit Function
    End If

    ' ·Ê ›ÌÂ "-" √Ê "/"
    If InStr(StrDate, "-") > 0 Or InStr(StrDate, "/") > 0 Then
        Dim sep As String
        If InStr(StrDate, "-") > 0 Then
            sep = "-"
        Else
            sep = "/"
        End If
        parts = Split(StrDate, sep)
        If UBound(parts) = 2 Then
            ' Õœœ «· — Ì»: ·Ê «·”‰… > 31 Ê > 12 ›ÂÌ ”‰… ›Ì «·√Ê·
            If val(parts(0)) > 31 Then
                ' yyyy-mm-dd √Ê yyyy/m/d
                YearStr = Format$(val(parts(0)), "0000")

                If IsNumeric(parts(1)) Then
                    MonthStr = Format$(val(parts(1)), "00")
                Else
                    MonthStr = MonthNameToNumber(parts(1))
                End If
                dayStr = Format$(val(parts(2)), "00")
            Else
                ' d/m/yyyy √Ê dd-mm-yyyy
                dayStr = Format$(val(parts(0)), "00")
                If IsNumeric(parts(1)) Then
                    MonthStr = Format$(val(parts(1)), "00")
                Else
                    MonthStr = MonthNameToNumber(parts(1))
                End If

                YearStr = Format$(val(parts(2)), "0000")
            End If
            ParseAnyDate = YearStr & "-" & MonthStr & "-" & dayStr
            Exit Function
        End If
    End If

    ' ·Ê «· «—ÌŒ Õ—Ê› („À· 3 „«ÌÊ 2025 √Ê 25 December 2024)
    parts = Split(StrDate, " ")
    If UBound(parts) = 2 Then
        dayStr = Format$(val(parts(0)), "00")
        MonthStr = MonthNameToNumber(parts(1)) ' «·œ«·… «··Ì  Õ  œÌ
        YearStr = Format$(val(parts(2)), "0000")
        ParseAnyDate = YearStr & "-" & MonthStr & "-" & dayStr
        Exit Function
    End If

    ' ·Ê ›‘· «· ÕÊÌ·° Ì—Ã⁄ «·‰’ «·√’·Ì
    ParseAnyDate = StrDate
End Function

Function MonthNameToNumber(ByVal m As String) As String
    m = LCase$(Trim(m))
    Select Case left(m, 3)
        Case "Ì‰«", "jan": MonthNameToNumber = "01"
        Case "›»—", "feb": MonthNameToNumber = "02"
        Case "„«—", "mar": MonthNameToNumber = "03"
        Case "√»—", "«»", "apr": MonthNameToNumber = "04"
        Case "„«Ì", "may": MonthNameToNumber = "05"
        Case "ÌÊ‰", "jun": MonthNameToNumber = "06"
        Case "ÌÊ·", "jul": MonthNameToNumber = "07"
        Case "√€”", "«€", "aug": MonthNameToNumber = "08"
        Case "”» ", "sep": MonthNameToNumber = "09"
        Case "√ﬂ ", "«ﬂ ", "oct": MonthNameToNumber = "10"
        Case "‰Ê›", "nov": MonthNameToNumber = "11"
        Case "œÌ”", "dec": MonthNameToNumber = "12"
        Case Else: MonthNameToNumber = "00"
    End Select
End Function

Private Sub SaveExcelFile2(FileName)
    On Error GoTo eh

    If FileName = "" Then
        MsgBox "«Œ — „·› «Ê·«"
        Exit Sub
    End If
   Dim mBranchID As Long
   Dim mIsStart As Boolean
Dim LoadExcelFlage As Boolean
Dim rsDummy As ADODB.Recordset
Dim s As String
Dim zatcaStatus As Integer
Dim mDateCh As Date
   Dim mBranchIDReSave As Integer
excelFileNameFullPath = FileName
Dim mTypeInvoice As Integer
    Dim i              As Long
    
    Dim mPrice         As Double
    Dim RsData         As New ADODB.Recordset
    Dim AllFinshedRows As Integer
    Dim allExcelRows   As Integer
    Dim startTime      As Date
    Dim moConn         As New ADODB.Connection
    Dim mrs            As ADODB.Recordset
    Dim tblname        As String
    'Dim shortFileName  As String
     '******************
    Dim POSname      As String
    Dim ItemName     As String
    Dim Qty          As String
    Dim paymethod    As String
    Dim SerialNo     As String
    Dim mDate        As String
    Dim mEmpName     As String
    
    Dim dstore       As Integer
    Dim dBox         As Integer
    Dim usertype     As Integer
            
    Dim userbranchid As Integer
    Dim CUSTID       As Integer
    
    Dim mVATValue       As Double
    Dim mEmpID2 As Long
    'GetBranchData branch_id, dstore, dBox
      Dim isFromExcel As Boolean
    
    'intDef
   
  
    '     Me.dcBranch.BoundText = userbranchid
    '     Me.DCboStoreName.BoundText = dstore

     
    Dim mItemCode As String
    ' lblTime.Visible = True
  '  Me.Enabled = False
    moConn.CursorLocation = adUseClient
    Dim rsCheck As New ADODB.Recordset
    Dim sfo     As New FileSystemObject

    'For i = 1 To grdFiles.rows - 1
    '        filename = grdFiles.TextMatrix(i, grdFiles.ColIndex("FileName"))
    shortFileName = sfo.GetFileName(FileName) 'grdFiles.TextMatrix(i, grdFiles.ColIndex("Name"))
    '
    moConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source='" & FileName & "'; Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
    Set mrs = moConn.OpenSchema(adSchemaTables)
   
    '****************
    grd(0).rows = 1
    Dim mMasterGui As String
    If Not mrs.EOF Then
        tblname = mrs.Fields("table_name").value
        RsData.CursorLocation = adUseClient
        RsData.Open "Select  *   from [" & tblname & "]", moConn, adOpenKeyset, adLockReadOnly
        mMasterGui = GenerateGUID
        s = "Select *, {" & mMasterGui & "} as  GroupUniqueFileMaster  from [" & tblname & "] "
        loadgrid s, grd(0), True, False, False, moConn
                           
        s = "Select * from tblEInvoice2 where id =0 "
        saveGrid s, grd(0), "InvoiceId", "serial"
        
        s = " update tblEInvoice2 set InvoiceTypeCodeID = '388',InvoiceTypeCodename ='0100000',Transaction_ID = invoiceid"
        s = s & " Where IsNull(VatValue,0) = 0 Or DefaultInvoicetype = 0 and GroupUniqueFileMaster = '" & Trim(mMasterGui) & "'"

        Cn.Execute s
            s = " update tblEInvoice2 set InvoiceTypeCodeID = '388',InvoiceTypeCodename ='0200000',Transaction_ID = invoiceid"
        s = s & " Where IsNull(VatValue,0) <> 0 and DefaultInvoicetype = 1 and GroupUniqueFileMaster = '" & Trim(mMasterGui) & "'"
    Cn.Execute s
    
    
    Dim sql As String

sql = ""
sql = sql & "IF OBJECT_ID('tempdb..#TmpGroups') IS NOT NULL DROP TABLE #TmpGroups;" & vbCrLf

sql = sql & "SELECT InvoiceID, NEWID() AS NewGroupCode " & vbCrLf
sql = sql & "INTO #TmpGroups " & vbCrLf
sql = sql & "FROM tblEInvoice2 " & vbCrLf
sql = sql & "WHERE GroupUniqueFileMaster = '" & Trim(mMasterGui) & "' " & vbCrLf
sql = sql & "GROUP BY InvoiceID;" & vbCrLf

sql = sql & "UPDATE t " & vbCrLf
sql = sql & "SET t.GroupUniqueCode = g.NewGroupCode " & vbCrLf
sql = sql & "FROM tblEInvoice2 t " & vbCrLf
sql = sql & "INNER JOIN #TmpGroups g ON t.InvoiceID = g.InvoiceID " & vbCrLf
sql = sql & "WHERE t.GroupUniqueFileMaster = '" & Trim(mMasterGui) & "';" & vbCrLf

' (???????) ??? ?????? ?????? ??????? ??? ???????
sql = sql & "DROP TABLE #TmpGroups;" & vbCrLf

' ??? ?????
Cn.Execute sql


    
s = ""
s = s & "SELECT " & vbCrLf
s = s & "    InvoiceID,GroupUniqueFileMaster,GroupUniqueCode," & vbCrLf
s = s & "    InvoiceID as Transaction_ID," & vbCrLf
s = s & "    Identificationid as Identificationid," & vbCrLf

s = s & "    DefaultInvoiceType," & vbCrLf
s = s & "    ExcelRow," & vbCrLf
s = s & "    ExcelFile," & vbCrLf
s = s & "    IssueDate," & vbCrLf
s = s & "    IssueTim," & vbCrLf
s = s & "    DocumentCurrencyCode," & vbCrLf
s = s & "    TaxCurrencyCode," & vbCrLf
s = s & "    StreetName," & vbCrLf
s = s & "    BuildingNumber," & vbCrLf
s = s & "    CityName," & vbCrLf
s = s & "    PostalZone," & vbCrLf
s = s & "    CitySubdivisionName," & vbCrLf
s = s & "    RegistrationName," & vbCrLf
s = s & "    CompanyID," & vbCrLf
s = s & "    CoCRCode," & vbCrLf
s = s & "    SUM(PayableAmount) AS PayableAmount," & vbCrLf
s = s & "    SUM(VatValue) AS VatValue," & vbCrLf
s = s & "    round(SUM(VatValue) /(SUM(Price*qty)   ) * 100,1)  AS TaxCategoryPercent ," & vbCrLf
'15' AS TaxCategoryPercent
s = s & "    Id700," & vbCrLf
s = s & "    QrCodeData," & vbCrLf
s = s & "    QrCodeDataPath," & vbCrLf
s = s & "    zatcaStatus," & vbCrLf
s = s & "    InvoiceTypeCodeID," & vbCrLf
s = s & "    InvoiceTypeCodename," & vbCrLf
s = s & "    AdditionalDocumentReferencePIH," & vbCrLf
s = s & "    InvoiceDocumentReferenceID," & vbCrLf
s = s & "    AdditionalDocumentReferenceICVUUID," & vbCrLf
s = s & "    ActualDeliveryDate," & vbCrLf
s = s & "    LatestDeliveryDate," & vbCrLf
s = s & "    RecTime " & vbCrLf
s = s & "FROM tblEInvoice2 " & vbCrLf
s = s & "GROUP BY " & vbCrLf
s = s & "    InvoiceID," & vbCrLf
s = s & "    DefaultInvoiceType," & vbCrLf
s = s & "    ExcelRow," & vbCrLf
s = s & "    ExcelFile," & vbCrLf
s = s & "    IssueDate," & vbCrLf
s = s & "    IssueTim," & vbCrLf
s = s & "    DocumentCurrencyCode," & vbCrLf
s = s & "    TaxCurrencyCode," & vbCrLf
s = s & "    StreetName," & vbCrLf
s = s & "    BuildingNumber," & vbCrLf
s = s & "    CityName," & vbCrLf
s = s & "    PostalZone," & vbCrLf
s = s & "    CitySubdivisionName," & vbCrLf
s = s & "    RegistrationName," & vbCrLf
s = s & "    CompanyID," & vbCrLf
s = s & "    CoCRCode," & vbCrLf
s = s & "    Id700," & vbCrLf
s = s & "    QrCodeData," & vbCrLf
s = s & "    QrCodeDataPath," & vbCrLf
s = s & "    zatcaStatus," & vbCrLf
s = s & "    InvoiceTypeCodeID," & vbCrLf
s = s & "    InvoiceTypeCodename," & vbCrLf
s = s & "    AdditionalDocumentReferencePIH," & vbCrLf
s = s & "    InvoiceDocumentReferenceID," & vbCrLf
s = s & "    AdditionalDocumentReferenceICVUUID," & vbCrLf
s = s & "    ActualDeliveryDate," & vbCrLf
s = s & "    LatestDeliveryDate,Identificationid,   "
   
s = s & "    RecTime;"

        loadgrid s, grd(0), True, False
        
        
         s = "Select * from tblEInvoice where id =0 "
        saveGrid s, grd(0), "InvoiceId", "serial"
        
        s = "Select * from tblEInvoice "
        Set rsDummy = New ADODB.Recordset
        
        rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
        Do While Not rsDummy.EOF
            
        SaveQRCode "tblEInvoice", "ID", val(rsDummy!ID & ""), val(rsDummy!invoiceID & ""), rsDummy!IssueDate & "", _
        val(rsDummy!PayableAmount & ""), Picture1, 0, val(rsDummy!VATValue & ""), val(rsDummy!PayableAmount & "")
            
            rsDummy.MoveNext
        Loop
                         

        
        Exit Sub
        Dim RowID          As Integer
          
        Dim isREcoredSaved As Boolean
        Dim strQuery       As String
        Dim OLDSec         As Long
        Dim Secondes       As Long

        Dim AllSec         As Long
        
        RowID = 0
        AllFinshedRows = 0
        Dim currentRows As Long
        Dim AllFileRows As Long
        RsData.MoveLast
        Dim LngFindRow As Long
        AllFileRows = RsData.RecordCount
        allExcelRows = AllFileRows
        RsData.MoveFirst
        'val(grdFiles.TextMatrix(i, grdFiles.ColIndex("Rows")))
        Dim RsStore As New ADODB.Recordset
        
        Do While Not RsData.EOF
        
           
           
'               DB_CreateField "tblEInvoice", "InvoicetID", adInteger, adColNullable, , , ""
'        DB_CreateField "tblEInvoice", "DefaultInvoicetype", adInteger, adColNullable, , , ""
'
'
'            DB_CreateField "tblEInvoice", "ExcelRow", adInteger, adColNullable, , , ""
'    DB_CreateField "tblEInvoice", "ExcelFile", adVarWChar, adColNullable, 400, , "", False, True, , True
'
'        DB_CreateField "tblEInvoice", "IssueDate", adDBTimeStamp, adColNullable, , , "", False, True
'        DB_CreateField "tblEInvoice", "IssueTim", adDBTimeStamp, adColNullable, , , "", False, True
'        DB_CreateField "tblEInvoice", "DocumentCurrencyCode", adVarWChar, adColNullable, 400, , "", False, True, , True
'        DB_CreateField "tblEInvoice", "TaxCurrencyCode", adVarWChar, adColNullable, 400, , "", False, True, , True
'        DB_CreateField "tblEInvoice", "StreetName", adVarWChar, adColNullable, 400, , "", False, True, , True
'        DB_CreateField "tblEInvoice", "BuildingNumber", adVarWChar, adColNullable, 400, , "", False, True, , True
'        DB_CreateField "tblEInvoice", "CityName", adVarWChar, adColNullable, 400, , "", False, True, , True
'        DB_CreateField "tblEInvoice", "PostalZone", adVarWChar, adColNullable, 400, , "", False, True, , True
'        DB_CreateField "tblEInvoice", "CitySubdivisionName", adVarWChar, adColNullable, 400, , "", False, True, , True
'        DB_CreateField "tblEInvoice", "RegistrationName", adVarWChar, adColNullable, 400, , "", False, True, , True
'        DB_CreateField "tblEInvoice", "CompanyID", adVarWChar, adColNullable, 400, , "", False, True, , True
'        DB_CreateField "tblEInvoice", "ItemName", adVarWChar, adColNullable, 400, , "", False, True, , True
'        DB_CreateField "tblEInvoice", "Qty", adDouble, adColNullable, , , "    ", False, True
'        DB_CreateField "tblEInvoice", "Price", adDouble, adColNullable, , , "    ", False, True
'        DB_CreateField "tblEInvoice", "CoCRCode", adVarWChar, adColNullable, 400, , "", False, True, , True
'
'        DB_CreateField "tblEInvoice", "PayableAmount", adDouble, adColNullable, , , "    ", False, True
'
'
'
'        DB_CreateField "tblEInvoice", "Id700", adVarWChar, adColNullable, 255, , "", False, True, , True
'        DB_CreateField "tblEInvoice", "serial", adInteger, adColNullable, , , ""
'
'
        
           
            mVATValue = val(RsData.Fields("VatValue"))
         '   POSname = RsData.Fields("‰ﬁÿ… »Ì⁄")
            Dim itemNameTmp As String
            itemNameTmp = RsData.Fields("ItemName") & ""
            
            Qty = RsData.Fields("Qty") & ""  '  Trim(Split(itemNameTmp, "x")(0))
            ItemName = itemNameTmp ' Trim(Split(itemNameTmp, "x")(1))
            paymethod = 1 ' RsData.Fields("ÿ—Ìﬁ… «·œ›⁄")
            SerialNo = RsData.Fields("ID") & ""
            mDate = RsData.Fields("IssueDate") & ""
           ' mEmpName = RsData.Fields("«”„⁄ «·»«∆⁄")
            
            mPrice = val(RsData.Fields("Price") & "")
            
         
            
         
            
            RowID = RowID + 1
            AllFinshedRows = AllFinshedRows + 1
            currentRows = currentRows + 1
            ' lbl(32).Caption = "F[" & currentRows & "]>[" & AllFileRows & "] A[" & AllFinshedRows & "]>[" & AllFileRows & "]"
            DoEvents
            strQuery = "SELECT Count(*) cnt "
            strQuery = strQuery & "From tblEInvoice "
            strQuery = strQuery & "WHERE ExcelFile = '" & shortFileName & "' "
            
            strQuery = strQuery & "  AND ExcelRow =  " & RowID & " ;"
            ' rsCheck.CursorLocation = adUseClient
            rsCheck.Open strQuery, Cn, adOpenForwardOnly, adLockReadOnly
            isREcoredSaved = rsCheck!cnt > 0
            rsCheck.Close

            '*********************
            If isREcoredSaved Then
               ' GoTo NextRow
            End If
            excelShortName = shortFileName
            excelRowNo = RowID
            startTime = Now
            '  btnpay_Click 0
            If Trim(SerialNo) <> "" Then
                'SaveItemsExcelMeth_New RsData, RowID, shortFileName
                '*************************
                '*************************
                
                ' XPDtbBill.value = Replace(Replace(Replace(mdate, "„", "PM"), "’", ""), "˛", "")
                Dim arr() As String
                arr = Split(ItemName, " ")
               
                s = ""
                s = s & "SELECT ItemCode,ItemID "
                s = s & "FROM TblItems "
                If UBound(arr) = 0 Then
                  '  s = s & "WHERE Fullcode LIKE N'%" & mItemCode & "%' "  'AND ItemName LIKE '%91%' "
                Else
                   ' s = s & "WHERE 1 = 1   "
                   
                    For i = 0 To UBound(arr)
                        If arr(i) <> "" Then
                    '        s = s & " And ItemName LIKE N'%" & Arr(i) & "%'  "
                        End If
                    Next
            
                End If
                 s = s & "WHERE Fullcode LIKE N'%" & mItemCode & "%' "  'AND ItemName LIKE '%91%' "
                Dim rsITem As New ADODB.Recordset
                Set rsITem = New ADODB.Recordset
                rsITem.Open s, Cn, adOpenForwardOnly, adLockReadOnly
                If rsITem.EOF Then
                    GoTo NextRow
                Else
                  
                    Dim dt As String
                    dt = RsData.Fields("business_date")
                    Dim mydt As Date
                    Dim mDay, mMonth, mYear, mhour, mmenut, mAmPm
                    Dim Arrt
                    Arrt = Split(dt, " ")
                    Dim arrDate() As String, arrTime
                    arrDate = Split(Arrt(0), "-")
                  '  mydt = DateSerial(val(GetNumbers(arrDate(2))), val(GetNumbers(arrDate(1))), val(GetNumbers(arrDate(0))))

                  
                '    TxtQuantity.text = Qty
                '    TxtPrice.text = mPrice
                 
                                    
                End If
                'CMDPAy_Click 2
                isFromExcel = True
                'If optsale(1).value = True Then 'return
                
                '  TxtNetValue.text = LblFinal.Caption
                '  TxtPayedValue2 = LblFinal.Caption
                '  TxtNetValue2 = LblFinal.Caption
                
                '      TxtPayedValue.text = LblFinal.Caption 'TxtNetValue.text
                'End If
               
            '    Cmd_Click (2)
                
            End If
       

            '*********************
NextRow:
RsData.MoveNext
     If Not RsData.EOF Then
        If Trim(SerialNo) <> (RsData.Fields("order_reference") & "") Then
            

            'SaveData
            ' btnNew_Click 0
            'End If

            OLDSec = AllSec
            Secondes = DateDiff("s", startTime, Now)

            AllSec = ((allExcelRows - AllFinshedRows) * Secondes)

            If AllSec = 0 Then
                AllSec = OLDSec
            End If

            
        Else
           
        End If
    Else
        
     End If
     
            '  End If
            
            
        Loop

        RsData.Close
    End If

    mrs.Close
    moConn.Close
    isFromExcel = False
    'Next

    Me.Enabled = True
    lbl(98).Visible = False
    'grdFiles.rows = 1
    MsgBox " „ Õ›Ÿ «·Õ—ﬂ« "
    Exit Sub
eh:
   
    
    MsgBox Err.Description
End Sub
Private Sub BtnPrint_Click(Index As Integer)
  
   print_report66 Index
End Sub
Function print_report66(Optional Ind As Integer = 0)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim mCusType As String
    Dim StrFileName As String
    Dim Msg As String
       Dim mSql1 As String
    Dim mSql2 As String
    mCusType = 1

    Dim X As Integer
'    MySQL = "Select * from TblAging  "
'
'     MySQL = MySQL & " WHERE 1 = 1  "
'    If Not IsNull(DTP_Date.value) Then
'        MySQL = MySQL & " and TblAging.DueDate <=" & SQLDate(DTP_Date.value, True) & ""
'    End If
'
'        If val(DCboItemsName2.BoundText) <> 0 Then
'            MySQL = MySQL & " and TblAging.CusID =" & val(DCboItemsName2.BoundText) & ""
'        End If
'
'    'StrCusID = ""
'    If CheckAllCustomer.value = vbChecked Then
'         If StrCusID.Text <> "" Then
'            MySQL = MySQL & " and TblAging.CusID in (" & (StrCusID.Text) & ")"
'        End If
'    End If
'     MySQL = MySQL & " ORDER BY DueDate  "
'    RsData.Open MySQL, Cn, adOpenKeyset, adLockReadOnly
'    If Not RsData.EOF Then
'          If SystemOptions.UserInterface = ArabicInterface Then
'                        X = MsgBox("ÌÊÃœ ⁄„— œÌ‰  „ ⁄„·Â „”»ﬁ« » «—ÌŒ " & DTP_Date.value & "" & " Â·  Êœ ⁄—÷Â ‰⁄„/·«", vbInformation + vbYesNo)
'                    Else
'                        X = MsgBox("No Contract For This Employee Create Contarct y / n", vbInformation + vbYesNo)
'                    End If
'
'                    If X = vbYes Then
'                        loadgrid MySQL, grdAging, True, False
'                        PrintAging Ind
'                        Exit Function
'                    End If
'    End If
'
'    RsData.Close
 
    

    
    grdAging2.rows = 1
    grdAging.rows = 1
    
Dim mWhereCus As String


'-------------------------------
   Dim mCusTypeStr As String
  
  mCusType = 1
   
MySQL = ""
'MySQL = MySQL & " SELECT "

   



MySQL = MySQL & "   SELECT XB.account_code,"
MySQL = MySQL & "          Xb.Transactionstypename,"
MySQL = MySQL & "          XB.duedate,"
MySQL = MySQL & "          diffdate,"
MySQL = MySQL & "          XB.notedate,"
MySQL = MySQL & "          XB.notetype,"
MySQL = MySQL & "          XB.note_value                                  TransNet,"
MySQL = MySQL & "          XB.cusname,"
MySQL = MySQL & "          xb.ageid,"
MySQL = MySQL & "          XB.NoteSerial1,"
MySQL = MySQL & "          dbo.ageng_type.NAME,"
MySQL = MySQL & "          dbo.ageng_type.[from],"
MySQL = MySQL & "          dbo.ageng_type.[to],"
MySQL = MySQL & "          dbo.ageng_type.color,"
MySQL = MySQL & "          dbo.ageng_type.namee,"
MySQL = MySQL & "          Isnull(Transactionstypename, 'ﬁÌœ «›  «ÕÏ') AS TransactionTypeName"
MySQL = MySQL & "   FROM   (SELECT dev.item_id"
MySQL = MySQL & "          AS"
MySQL = MySQL & "          account_code,"
MySQL = MySQL & "          transactiontypes.StockEffect"
MySQL = MySQL & "          AS credit_or_debit,"
MySQL = MySQL & "          transactions.BranchID"
MySQL = MySQL & "          AS branch_id,"
MySQL = MySQL & "          dev.Transaction_ID"
MySQL = MySQL & "          AS Transactions_id,"
MySQL = MySQL & "          transactiontypes.transactiontypename"
MySQL = MySQL & "          AS Transactionstypename,"
MySQL = MySQL & "          transactions.Transaction_Date"
MySQL = MySQL & "          AS DueDate,"
MySQL = MySQL & "          transactions.Transaction_Date"
MySQL = MySQL & "          AS notedate,"
MySQL = MySQL & "          transactions.Transaction_Type"
MySQL = MySQL & "          AS notetype,"
MySQL = MySQL & "          transactions.NoteSerial1"
MySQL = MySQL & "          AS NoteSerial1,"
MySQL = MySQL & "          Isnull(dev.ShowQty, 0)"
MySQL = MySQL & "          AS Note_Value,"
MySQL = MySQL & "          a.ItemName"
MySQL = MySQL & "          AS CusName,"
MySQL = MySQL & "          Isnull(dev.ShowQty , 0)"
MySQL = MySQL & "          AS TransNet,"
MySQL = MySQL & "          dbo.Getdeptageid(Datediff(day, transactions.transaction_date,"
MySQL = MySQL & "                           " & SQLDate(DTP_Date.value, True) & " )) AS"
MySQL = MySQL & "          AgeID,"
MySQL = MySQL & "          Datediff(day, transactions.transaction_date, " & SQLDate(DTP_Date.value, True) & " )"
MySQL = MySQL & "          AS DiffDate"
MySQL = MySQL & "          FROM   transaction_details AS dev"
MySQL = MySQL & "                 INNER JOIN transactions"
MySQL = MySQL & "                         ON transactions.transaction_id = dev.transaction_id"
MySQL = MySQL & "                 INNER JOIN transactiontypes"
MySQL = MySQL & "                         ON transactions.transaction_type ="
MySQL = MySQL & "                            transactiontypes.Transaction_Type"
MySQL = MySQL & "                 LEFT OUTER JOIN tblitems AS a"
MySQL = MySQL & "                              ON a.itemid = dev.item_id"
MySQL = MySQL & "          Where (IsNull(dev.ShowQty , 0) <> 0)"
MySQL = MySQL & "                 AND ( transactiontypes.stockeffect = 1 )"
MySQL = MySQL & "                 and transactions.transaction_type <> 3"

MySQL = MySQL & "                  Union all"
     
     
     
     
MySQL = MySQL & "               SELECT dev.item_id"
MySQL = MySQL & "                 AS"
MySQL = MySQL & "                 account_code,"
MySQL = MySQL & "                 transactiontypes.StockEffect"
MySQL = MySQL & "                 AS credit_or_debit,"
MySQL = MySQL & "                 transactions.BranchID"
MySQL = MySQL & "                 AS branch_id,"
MySQL = MySQL & "                 dev.Transaction_ID"
MySQL = MySQL & "                 AS Transactions_id,"
MySQL = MySQL & "                 transactiontypes.transactiontypename"
MySQL = MySQL & "                 AS Transactionstypename,"
MySQL = MySQL & "                 transactions.Transaction_Date"
MySQL = MySQL & "                 AS DueDate,"
MySQL = MySQL & "                 transactions.Transaction_Date"
MySQL = MySQL & "                 AS notedate,"
MySQL = MySQL & "                 transactions.Transaction_Type"
MySQL = MySQL & "                 AS notetype,"
MySQL = MySQL & "                 transactions.NoteSerial1"
MySQL = MySQL & "                 AS NoteSerial1,"
MySQL = MySQL & "                 Isnull(dev.ShowQty , 0)"
MySQL = MySQL & "                 AS Note_Value,"
MySQL = MySQL & "                 a.ItemName"
MySQL = MySQL & "                 AS CusName,"
MySQL = MySQL & "                 Isnull(dev.ShowQty , 0)"
MySQL = MySQL & "                 AS TransNet,"
MySQL = MySQL & "                 dbo.Getdeptageid(Datediff(day, transactions.transaction_date,"
MySQL = MySQL & "                                  " & SQLDate(DTP_Date.value, True) & " )) AS"
MySQL = MySQL & "                 AgeID,"
MySQL = MySQL & "                 Datediff(day, transactions.transaction_date, " & SQLDate(DTP_Date.value, True) & " )"
MySQL = MySQL & "                 AS DiffDate"
MySQL = MySQL & "          FROM   transaction_details AS dev"
MySQL = MySQL & "                 INNER JOIN transactions"
MySQL = MySQL & "                         ON transactions.transaction_id = dev.transaction_id"
MySQL = MySQL & "                 INNER JOIN transactiontypes"
MySQL = MySQL & "                         ON transactions.transaction_type ="
MySQL = MySQL & "                            transactiontypes.Transaction_Type"
MySQL = MySQL & "                 LEFT OUTER JOIN tblitems AS a"
MySQL = MySQL & "                              ON a.itemid = dev.item_id"
MySQL = MySQL & "          Where (IsNull(dev.ShowQty, 0) <> 0)"
MySQL = MySQL & "                 AND ( transactiontypes.stockeffect = 1 )"
MySQL = MySQL & "                 and transactions.transaction_type = 3"
MySQL = MySQL & "               ) XB"
MySQL = MySQL & "                 RIGHT OUTER JOIN dbo.ageng_type"
MySQL = MySQL & "                               ON XB.ageid = dbo.ageng_type.id"
MySQL = MySQL & "          Where 1 = 1"
MySQL = MySQL & "                 AND XB.duedate <= " & SQLDate(DTP_Date.value, True) & " "
     
     
If val(DCboItemsName2.BoundText) <> 0 Then
        MySQL = MySQL & " and Account_Code  In (Select tblitems.ItemID from tblitems Where  tblitems.ItemID= " & val(DCboItemsName2.BoundText) & " )"
        
End If
        
    If CheckAllCustomer.value = vbChecked Then
         If StrCusID.Text <> "" Then
           ' MySQL = MySQL & " and TblCustemers.CusID in (" & (StrCusID.Text) & ")"
            MySQL = MySQL & " and Account_Code  In ( Select tblitems.ItemID from tblitems Where  tblitems.ItemID  in (" & (StrCusID.Text) & ") )"
        End If
    End If
        
        
MySQL = MySQL & "          ORDER  BY account_code,"
MySQL = MySQL & "                    XB.NoteSerial1,"
MySQL = MySQL & "                    Xb.dueDate"

mSql1 = MySQL
          



MySQL = ""


MySQL = ""
'MySQL = MySQL & " SELECT "

   



MySQL = MySQL & "   SELECT XB.account_code,"
MySQL = MySQL & "          Xb.Transactionstypename,"
MySQL = MySQL & "          XB.duedate,"
MySQL = MySQL & "          diffdate,"
MySQL = MySQL & "          XB.notedate,"
MySQL = MySQL & "          XB.notetype,"
MySQL = MySQL & "          XB.note_value                                  TransNet,"
MySQL = MySQL & "          XB.cusname,"
MySQL = MySQL & "          xb.ageid,"
MySQL = MySQL & "          XB.NoteSerial1,"
MySQL = MySQL & "          dbo.ageng_type.NAME,"
MySQL = MySQL & "          dbo.ageng_type.[from],"
MySQL = MySQL & "          dbo.ageng_type.[to],"
MySQL = MySQL & "          dbo.ageng_type.color,"
MySQL = MySQL & "          dbo.ageng_type.namee,"
MySQL = MySQL & "          Isnull(Transactionstypename, 'ﬁÌœ «›  «ÕÏ') AS TransactionTypeName"
MySQL = MySQL & "   FROM   (SELECT dev.item_id"
MySQL = MySQL & "          AS"
MySQL = MySQL & "          account_code,"
MySQL = MySQL & "          transactiontypes.StockEffect"
MySQL = MySQL & "          AS credit_or_debit,"
MySQL = MySQL & "          transactions.BranchID"
MySQL = MySQL & "          AS branch_id,"
MySQL = MySQL & "          dev.Transaction_ID"
MySQL = MySQL & "          AS Transactions_id,"
MySQL = MySQL & "          transactiontypes.transactiontypename"
MySQL = MySQL & "          AS Transactionstypename,"
MySQL = MySQL & "          transactions.Transaction_Date"
MySQL = MySQL & "          AS DueDate,"
MySQL = MySQL & "          transactions.Transaction_Date"
MySQL = MySQL & "          AS notedate,"
MySQL = MySQL & "          transactions.Transaction_Type"
MySQL = MySQL & "          AS notetype,"
MySQL = MySQL & "          transactions.NoteSerial1"
MySQL = MySQL & "          AS NoteSerial1,"
MySQL = MySQL & "          Isnull(dev.ShowQty , 0)"
MySQL = MySQL & "          AS Note_Value,"
MySQL = MySQL & "          a.ItemName"
MySQL = MySQL & "          AS CusName,"
MySQL = MySQL & "          Isnull(dev.ShowQty , 0)"
MySQL = MySQL & "          AS TransNet,"
MySQL = MySQL & "          dbo.Getdeptageid(Datediff(day, transactions.transaction_date,"
MySQL = MySQL & "                           " & SQLDate(DTP_Date.value, True) & " )) AS"
MySQL = MySQL & "          AgeID,"
MySQL = MySQL & "          Datediff(day, transactions.transaction_date, " & SQLDate(DTP_Date.value, True) & " )"
MySQL = MySQL & "          AS DiffDate"
MySQL = MySQL & "          FROM   transaction_details AS dev"
MySQL = MySQL & "                 INNER JOIN transactions"
MySQL = MySQL & "                         ON transactions.transaction_id = dev.transaction_id"
MySQL = MySQL & "                 INNER JOIN transactiontypes"
MySQL = MySQL & "                         ON transactions.transaction_type ="
MySQL = MySQL & "                            transactiontypes.Transaction_Type"
MySQL = MySQL & "                 LEFT OUTER JOIN tblitems AS a"
MySQL = MySQL & "                              ON a.itemid = dev.item_id"
MySQL = MySQL & "          Where (IsNull(dev.ShowQty , 0) <> 0)"
MySQL = MySQL & "                 AND ( transactiontypes.stockeffect = -1 )"
MySQL = MySQL & "                 and transactions.transaction_type <> 3"

MySQL = MySQL & "                  Union all"
     
     
     
     
MySQL = MySQL & "               SELECT dev.item_id"
MySQL = MySQL & "                 AS"
MySQL = MySQL & "                 account_code,"
MySQL = MySQL & "                 transactiontypes.StockEffect"
MySQL = MySQL & "                 AS credit_or_debit,"
MySQL = MySQL & "                 transactions.BranchID"
MySQL = MySQL & "                 AS branch_id,"
MySQL = MySQL & "                 dev.Transaction_ID"
MySQL = MySQL & "                 AS Transactions_id,"
MySQL = MySQL & "                 transactiontypes.transactiontypename"
MySQL = MySQL & "                 AS Transactionstypename,"
MySQL = MySQL & "                 transactions.Transaction_Date"
MySQL = MySQL & "                 AS DueDate,"
MySQL = MySQL & "                 transactions.Transaction_Date"
MySQL = MySQL & "                 AS notedate,"
MySQL = MySQL & "                 transactions.Transaction_Type"
MySQL = MySQL & "                 AS notetype,"
MySQL = MySQL & "                 transactions.NoteSerial1"
MySQL = MySQL & "                 AS NoteSerial1,"
MySQL = MySQL & "                 Isnull(dev.ShowQty , 0)"
MySQL = MySQL & "                 AS Note_Value,"
MySQL = MySQL & "                 a.ItemName"
MySQL = MySQL & "                 AS CusName,"
MySQL = MySQL & "                 Isnull(dev.ShowQty , 0)"
MySQL = MySQL & "                 AS TransNet,"
MySQL = MySQL & "                 dbo.Getdeptageid(Datediff(day, transactions.transaction_date,"
MySQL = MySQL & "                                  " & SQLDate(DTP_Date.value, True) & " )) AS"
MySQL = MySQL & "                 AgeID,"
MySQL = MySQL & "                 Datediff(day, transactions.transaction_date, " & SQLDate(DTP_Date.value, True) & " )"
MySQL = MySQL & "                 AS DiffDate"
MySQL = MySQL & "          FROM   transaction_details AS dev"
MySQL = MySQL & "                 INNER JOIN transactions"
MySQL = MySQL & "                         ON transactions.transaction_id = dev.transaction_id"
MySQL = MySQL & "                 INNER JOIN transactiontypes"
MySQL = MySQL & "                         ON transactions.transaction_type ="
MySQL = MySQL & "                            transactiontypes.Transaction_Type"
MySQL = MySQL & "                 LEFT OUTER JOIN tblitems AS a"
MySQL = MySQL & "                              ON a.itemid = dev.item_id"
MySQL = MySQL & "          Where (IsNull(dev.ShowQty , 0) <> 0)"
MySQL = MySQL & "                 AND ( transactiontypes.stockeffect = -1 )"
MySQL = MySQL & "                 and transactions.transaction_type = 3"
MySQL = MySQL & "               ) XB"
MySQL = MySQL & "                 RIGHT OUTER JOIN dbo.ageng_type"
MySQL = MySQL & "                               ON XB.ageid = dbo.ageng_type.id"
MySQL = MySQL & "          Where 1 = 1"
MySQL = MySQL & "                 AND XB.duedate <= " & SQLDate(DTP_Date.value, True) & " "
     
If val(DCboItemsName2.BoundText) <> 0 Then
        MySQL = MySQL & " and Account_Code  In (Select tblitems.ItemID from tblitems Where  tblitems.ItemID= " & val(DCboItemsName2.BoundText) & " )"
        
End If

    If CheckAllCustomer.value = vbChecked Then
         If StrCusID.Text <> "" Then
           ' MySQL = MySQL & " and TblCustemers.CusID in (" & (StrCusID.Text) & ")"
            MySQL = MySQL & " and Account_Code  In ( Select tblitems.ItemID from tblitems Where  tblitems.ItemID  in (" & (StrCusID.Text) & ") )"
        End If
    End If

MySQL = MySQL & "          ORDER  BY account_code,"
MySQL = MySQL & "                    XB.NoteSerial1,"
MySQL = MySQL & "                    Xb.dueDate"


          


'MySQL = MySQL & " ,AgeID"

   
   
   mSql2 = MySQL

    
        loadgrid mSql1, grdAging, True, False
        loadgrid mSql2, grdAging2, False, False
   
    
    
'
'
'    MySQL = MySQL & "                                          AND (DATEDIFF(DAY, '31-Aug-2020', RptLedger_Sub2.RecordDate) < 0)"
'    MySQL = MySQL & "                            ) XB"
'
'    MySQL = MySQL & "                               LEFT OUTER JOIN dbo.Ageng_type"
'    MySQL = MySQL & "                                    ON  XB.ID = dbo.Ageng_type.id"
'
'    MySQL = MySQL & "                        WHERE  XB.DueDate >= '01-Aug-2020'"
'    MySQL = MySQL & "                               AND XB.DueDate <= '31-Aug-2020'"
'
'    MySQL = MySQL & "                        Order By "
'    MySQL = MySQL & "                               Xb.ID , DueDate "
'
'
   
Dim s As String
   Dim i As Long
   Dim mValue As Double
   Dim mCusId As Long
    Dim j As Long
    Dim mValue2 As Double
   Dim mCusId2 As Long
    Dim mPayedValue As Double

Dim mAccount_Code As String
Dim mAccount_Code2 As String
Dim Balance As String

'If grdAging.Rows > 1 Then
'    mAccount_Code = Trim(grdAging.TextMatrix(1, grdAging.ColIndex("Account_Code")))
'    WriteCustomerBalPublic mAccount_Code, Balance, , 0, , , , , FromDate1.value, 1
'    grdAging.TextMatrix(1, grdAging.ColIndex("Balance")) = Balance
'End If

'If grdAging2.Rows > 1 Then
'    Balance = ""
'    mAccount_Code = Trim(grdAging2.TextMatrix(1, grdAging2.ColIndex("Account_Code")))
'    WriteCustomerBalPublic mAccount_Code, Balance, , 1, , , , , FromDate1.value, 1
'    grdAging2.TextMatrix(1, grdAging2.ColIndex("Balance")) = Balance
'End If
txtTotalStill = ""
Dim mJ As Long
mJ = 1
   For i = 1 To grdAging.rows - 1
        Label6(1).Caption = i
     
'     If I = 1 Then
'        mAccount_Code = Trim(grdAging.TextMatrix(I, grdAging.ColIndex("Account_Code")))
'        WriteCustomerBalPublic mAccount_Code, Balance, , 0, , , , , FromDate1.value, 1
'        grdAging.TextMatrix(I, grdAging.ColIndex("Balance")) = Balance
'     End If
      mValue = val(grdAging.TextMatrix(i, grdAging.ColIndex("TransNet")))
      mAccount_Code = Trim(grdAging.TextMatrix(i, grdAging.ColIndex("Account_Code")))
      
        If val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue"))) <> mValue Then
        
        
        mJ = grdAging2.FindRow(mAccount_Code, grdAging2.FixedRows, grdAging2.ColIndex("Account_Code"), False, True)
        'mJ = grdAging2.FindRow("dsfdsf", grdAging2.FixedRows, grdAging2.ColIndex("Account_Code"), False, True)
       ' For j = mJ To grdAging2.Rows - 1
'
       j = mJ
            If mJ <> -1 Then
                mValue2 = val(grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")))
                mAccount_Code2 = Trim(grdAging2.TextMatrix(j, grdAging2.ColIndex("Account_Code")))
                mPayedValue = val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")))
    
    
               If mValue2 <> 0 And mAccount_Code2 = mAccount_Code And mValue <> mPayedValue Then
    
                    If mValue - mPayedValue = mValue2 Then
                        grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) = mValue2
                        grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")) = 0
                    ElseIf mValue - mPayedValue > mValue2 Then
                        grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) = val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue"))) + mValue2
                        grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")) = 0
                    ElseIf mValue - mPayedValue < mValue2 Then
                        grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) = mPayedValue + mValue - mPayedValue
                        grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet")) = mValue2 - (mValue - mPayedValue)
                        grdAging.TextMatrix(i, grdAging.ColIndex("TransNetGrid2")) = mValue2 - (mValue - mPayedValue)
                        'mValue - grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")) + mValue2
                    End If
               End If
               grdAging2.TextMatrix(j, grdAging2.ColIndex("StillAmount")) = val(grdAging2.TextMatrix(j, grdAging2.ColIndex("TransNet"))) - val(grdAging2.TextMatrix(j, grdAging2.ColIndex("PayedValue")))
            End If
'
'
'            If mAccount_Code2 <> mAccount_Code Then
'                GoTo ExitFor
'            End If
'        Next
        
      End If
      'mJ = j + 1
ExitFor:
      grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")) = val(grdAging.TextMatrix(i, grdAging.ColIndex("TransNet"))) - val(grdAging.TextMatrix(i, grdAging.ColIndex("PayedValue")))
      If val(grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount"))) = 0 Then
        grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")) = ""
        grdAging.RowHidden(i) = True
      End If
      txtTotalStill = val(txtTotalStill) + val(grdAging.TextMatrix(i, grdAging.ColIndex("StillAmount")))
   Next
    
    s = "Delete TblAging "
    Cn.Execute s
    
    
    

    
    s = "Select * from TblAging  Where  AGEID = -10 "
    
    
    
    saveGrid s, grdAging, "StillAmount", "", "Credit_Or_Debit", 0



    Dim rsDummyT As New ADODB.Recordset
    Dim rsDummyT2 As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    s = " Select Account_Code,AGEID,CusName from TblAging"
    s = s & " GROUP BY Account_Code,AGEID,CusName"
    Set rsDummyT = New ADODB.Recordset
    rsDummyT.Open s, Cn, adOpenStatic, adLockReadOnly
    Dim ii As Long
    Do While Not rsDummyT.EOF
        s = "Select * from Ageng_type where Id Not In (Select  AGEID from TblAging Where  Account_Code = N'" & Trim(rsDummyT!Account_code & "") & "' )"
        Set rsDummyT2 = New ADODB.Recordset
        rsDummyT2.Open s, Cn, adOpenStatic, adLockReadOnly
        ii = ii + 1
        Label6(1).Caption = ii
        Do While Not rsDummyT2.EOF
            s = "Select AGEID,CusName ,[To] ,[from] ,Name,Account_code from TblAging where AGEID = -10 "
            rs.Open s, Cn, adOpenKeyset, adLockOptimistic
            rs.AddNew
            rs!ageid = rsDummyT2!ID
            rs!Account_code = rsDummyT!Account_code & ""
            rs!CusName = rsDummyT!CusName & ""
            rs!To = rsDummyT2!To & ""
            rs!From = rsDummyT2!From & ""
            If SystemOptions.UserInterface = ArabicInterface Then
                rs!Name = rsDummyT2!Name & ""
            Else
                rs!Name = rsDummyT2!Name & ""
            End If
            rs.update
            rs.Close
            rsDummyT2.MoveNext
        Loop
        's = "Select Account_Code,AGEID from TblAging Where  Account_Code = " & Trim(rsDummyT!Account_code & "")
        
        rsDummyT.MoveNext
    Loop
    
 
    s = "Select * from TblAging "
'    saveGrid s, grdAging2, "CusId", "Id", "Credit_Or_Debit", 1


    Set RsData = New ADODB.Recordset
    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "No data"
        End If
        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If
StrCusID = ""
    RsData.Close
    Set RsData = Nothing
    PrintAging Ind
    Screen.MousePointer = vbDefault
End Function


Private Sub PrintAging(ByVal Ind As Long)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
        
        
    
    MySQL = "Select * from TblAging  Order By AGEID"
   
 
        MySQL = " Select TblAging.*,a.Account_Serial,ta.aqarNo,ta.aqarname"
        MySQL = MySQL & " from TblAging LEFT OUTER JOIN ACCOUNTS AS a ON a.Account_Code = TblAging.Account_Code LEFT OUTER JOIN TblCustemers AS tc ON tc.Account_Code = a.Account_Code"
        MySQL = MySQL & " LEFT OUTER JOIN TblAqar AS ta ON tc.CusID = ta.ownerid"
        MySQL = MySQL & " Order By TblAging.AGEID"
            
    If Ind = 0 Then
    
        
            MySQL = " Select TblAging.*,a.Account_Serial,ta.aqarNo,ta.aqarname"
        MySQL = MySQL & " from TblAging LEFT OUTER JOIN ACCOUNTS AS a ON a.Account_Code = TblAging.Account_Code LEFT OUTER JOIN TblCustemers AS tc ON tc.Account_Code = a.Account_Code"
        MySQL = MySQL & " LEFT OUTER JOIN TblAqar AS ta ON tc.CusID = ta.ownerid "
        MySQL = MySQL & " WHERE ISNULL(StillAmount,0) <> 0"
        MySQL = MySQL & " Order By TblAging.AGEID"
        
        MySQL = " SELECT tblaging.*,"
        MySQL = MySQL & " a.ItemCode as   account_serial,"
        MySQL = MySQL & " aqarno = '',"
        MySQL = MySQL & " aqarname = ''"
        MySQL = MySQL & " From tblaging"
        MySQL = MySQL & "                LEFT OUTER JOIN TblItems AS a"
        MySQL = MySQL & "                      ON a.itemid= tblaging.account_code"
        
        
        MySQL = MySQL & "         Where IsNull(stillamount, 0) <> 0"
        MySQL = MySQL & "  ORDER  BY tblaging.ageid"

     '   MySQL = "SELECT * FROM TblAging WHERE ISNULL(StillAmount,0) <> 0"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "AgingItem1.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "AgingItem1.rpt"
        End If
    Else
        MySQL = " SELECT TblAging.* FROM TblAging INNER JOIN Ageng_type ON Ageng_type.id = TblAging.AGEID"
        MySQL = MySQL & " ORDER BY   Ageng_type.id,TblAging.Account_Code"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Aging2Item1.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Aging2Item1.rpt"
        End If
    End If
    
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "No data"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub Check1_Click()
 
  Dim i As Integer

    If Check1.value = vbChecked Then

        With Me.grd(0)
 
            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("ch")) = True
            Next i

        End With

    Else

        With Me.grd(0)

            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("ch")) = False
            Next i

        End With

    End If
    
       '     Me.lbl(14).Caption = val(Calculate_TotalSelected2)


End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index
        Case 0
        GetData
        Case 1
             Unload Me
        Case 2, 4
            Unload Me
        Case 3
          If optNetWork(10).value = True Then
                GetDataTrans3
            Else
                GetDataNetwork
            End If
        Case 5
        If Optrans(1).value = True Then
        GetDataTrans2
        Else
        GetDataTrans
        End If
        
            
            
    End Select

End Sub

Function GetResultsOld()
Dim s As String
Dim rsDummy As New ADODB.Recordset


Dim mCount1 As Double
Dim mCount2 As Double
    
    
Dim mCount3 As Double
Dim mCount4 As Double
    
    
    
Dim mCount5 As Double
Dim mCount6 As Double
Dim mCount7 As Double
Dim mCount8 As Double
If SystemOptions.CanUploadZakatOpt Then
    ConectionFirst
End If
    

    
    
    
    s = " SELECT count(Transaction_ID) as noofinvoices  "
      s = s & "          From dbo.transactions"
 

     
 
  s = s & "  Where (dbo.transactions.Transaction_Type = 21 or dbo.transactions.Transaction_Type = 9)"
                        
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    If SystemOptions.ZacatHandW Then
        s = s & " and dbo.Transactions.Transaction_Type=854798 "
    End If
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & " "
    End If
    
       
    If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     Transactions.BranchId In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & val(DcBranches(2).BoundText) & "))"
  End If
  
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and     Transactions.BranchId = " & val(DcBranches(3).BoundText)
        
        
  End If

  
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount1 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
    
       s = " SELECT count(NoteID) as noofinvoices  "
      s = s & "          From dbo.Notes"
 

     
 
  s = s & "  Where (dbo.Notes.NoteType = 9083 or dbo.Notes.NoteType =9082)"
                        
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "

    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.Notes.NoteDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.Notes.NoteDate  <=" & SQLDate(ToDate, True) & " "
    End If
    
    
        If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     notes.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
       
    
            
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       notes.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If

    
    
    
    
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount5 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
    
        s = " SELECT count(NoteID) as noofinvoices  "
      s = s & "          From dbo.notes_all"
 

     
 
  s = s & "  Where (dbo.notes_all.NoteType = 85 or dbo.notes_all.NoteType =85)"
                        
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.notes_all.NoteDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.notes_all.NoteDate  <=" & SQLDate(ToDate, True) & " "
    End If
    
                If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     notes_all.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
       
    
            
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       notes_all.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount7 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
    s = " SELECT count(ID) as noofinvoices  "
      s = s & "          From dbo.project_billl"
 

     
 
  s = s & "  Where 1 = 1 and project_billl.bill_type = 0"
                        
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.project_billl.bill_date >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.project_billl.bill_date <=" & SQLDate(ToDate, True) & " "
    End If
       If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     project_billl.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and     project_billl.branch_no = " & val(DcBranches(3).BoundText)
        
        
  End If
        
        
        
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
        mCount2 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    End If
    rsDummy.Close
    
    
       s = " SELECT count(InvoiceID) as noofinvoices  "
      s = s & "          From dbo.tblEInvoice"
 

     
 
  s = s & "  Where 1 = 1 "
                        
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate <=" & SQLDate(ToDate, True) & " "
    End If
    
      If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     tblEInvoice.branch_id In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
  
  
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       tblEInvoice.branch_id = " & val(DcBranches(3).BoundText)
        
        
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
        mCount8 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    End If
    rsDummy.Close
    Dim mCount111 As Integer
    mCount111 = 0
If SystemOptions.CanUploadZakatOpt Then
        
        
        
        
        s = " SELECT COUNT(PropertyDueBatchDetail.id) AS noofinvoices "
        s = s & " FROM dbo.PropertyDueBatchDetail "
         
            s = s & " INNER JOIN PropertyContractBatch ON PropertyContractBatch.Id = PropertyDueBatchDetail.PropertyContractBatchId "
            s = s & " INNER JOIN PropertyContract ON PropertyContract.Id = PropertyContractBatch.MainDocId "
        s = s & " WHERE 1 = 1 "
        
         '   s = s & "     and PropertyContract.id not in (Select PropertyContractTermination.PropertyContractId from PropertyContractTermination)"

            If Not IsNull(FromDate.value) Then
                s = s & " AND PropertyContractBatch.BatchDate >= " & SQLDate(FromDate, True) & " "
            End If
            s = s & " AND ISNULL(IsSelected, 0) = 1"
            If Not IsNull(ToDate.value) Then
                s = s & " AND PropertyContractBatch.BatchDate <= " & SQLDate(ToDate, True) & " "
            End If
            
            ' ›· — Õ«·… ZATCA
            
        
        Set rsDummy = New ADODB.Recordset
        
            
                  '
        rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
        If rsDummy.RecordCount > 0 Then
            mCount111 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
        End If
        
        rsDummy.Close
    End If
    
    
    Label18.Caption = mCount2 + mCount1 + mCount5 + mCount7 + mCount8 + mCount111
     
       s = " SELECT count(Transaction_ID) as noofinvoices  "
      s = s & "          From dbo.transactions"
 

     
 
  s = s & "  Where (dbo.transactions.Transaction_Type = 21 or dbo.transactions.Transaction_Type = 9)"
                        
  s = s & "   and isnull( Transactions.zatcaStatus,0)=1   "
        If SystemOptions.ZacatHandW Then
        s = s & " and dbo.Transactions.Transaction_Type=854798 "
    End If
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & " "
    End If
       
    If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     Transactions.BranchId In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
         
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and     Transactions.BranchId = " & val(DcBranches(3).BoundText)
        
        
  End If
         
  
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
     mCount3 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    End If
    rsDummy.Close
    
    
      s = " SELECT count(ID) as noofinvoices  "
      s = s & "          From dbo.project_billl"
 

     
 
  s = s & "  Where 1 = 1 and project_billl.bill_type = 0"
                        
    s = s & "   and isnull( project_billl.zatcaStatus,0)=1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.project_billl.bill_date >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.project_billl.bill_date <=" & SQLDate(ToDate, True) & " "
    End If
       If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     project_billl.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
  
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and     project_billl.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If
         

        
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
        mCount4 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    End If
    rsDummy.Close
    
    
    
         s = " SELECT count(NoteID) as noofinvoices  "
      s = s & "          From dbo.notes_all"
 

     
 
  s = s & "  Where (dbo.notes_all.NoteType = 85 or dbo.notes_all.NoteType =85)"
s = s & "   and isnull( notes_all.zatcaStatus,0)=1   "
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.notes_all.NoteDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.notes_all.NoteDate  <=" & SQLDate(ToDate, True) & " "
    End If
    s = s & "   and isnull( notes_all.zatcaStatus,0)=1   "
    
If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     notes_all.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
       
    
            
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       notes_all.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount7 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
    
      
       s = " SELECT count(NoteID) as noofinvoices  "
      s = s & "          From dbo.Notes"
 

     
 
  s = s & "  Where (dbo.Notes.NoteType = 9083 or dbo.Notes.NoteType =9082)"
s = s & "   and isnull( Notes.zatcaStatus,0)=1   "
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.Notes.NoteDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.Notes.NoteDate  <=" & SQLDate(ToDate, True) & " "
    End If
    s = s & "   and isnull( Notes.zatcaStatus,0)=1   "
    
                If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     notes.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
       
    
            
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       notes.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount6 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
     
    
    
    
          s = " SELECT count(InvoiceID) as noofinvoices  "
      s = s & "          From dbo.tblEInvoice"
 

     
 
  s = s & "  Where 1 = 1 "
s = s & "   and isnull( tblEInvoice.zatcaStatus,0)=1   "
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate  <=" & SQLDate(ToDate, True) & " "
    End If
    
      If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     tblEInvoice.branch_id In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
  
  
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       tblEInvoice.branch_id = " & val(DcBranches(3).BoundText)
        
        
    End If
    s = s & "   and isnull( tblEInvoice.zatcaStatus,0)=1   "
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount8 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
     
     
     
   Dim mCount9 As Double
     
       
    
          s = " SELECT count(id) as noofinvoices  "
      s = s & "          From dbo.TblHandWages"
 

     
 
  s = s & "  Where 1 = 1 "
s = s & "   and isnull( TblHandWages.zatcaStatus,0)=1   "
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.TblHandWages.recordDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.TblHandWages.recordDate  <=" & SQLDate(ToDate, True) & " "
    End If
    s = s & "   and isnull( TblHandWages.zatcaStatus,0)=1   "
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount9 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
    
    
    Dim mCount10 As Double

s = " SELECT COUNT(id) AS noofinvoices "
s = s & " FROM dbo.tblContractInsAllocationsDetails "

s = s & " WHERE 1 = 1 "
s = s & " AND ISNULL(tblContractInsAllocationsDetails.zatcaStatus, 0) = 1 "

If Not IsNull(FromDate.value) Then
    s = s & " AND tblContractInsAllocationsDetails.Installdate >= " & SQLDate(FromDate, True) & " "
End If

If Not IsNull(ToDate.value) Then
    s = s & " AND tblContractInsAllocationsDetails.Installdate <= " & SQLDate(ToDate, True) & " "
End If

Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly

If rsDummy.RecordCount > 0 Then
    mCount10 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
End If

rsDummy.Close


   Dim mCount11 As Double
 If SystemOptions.CanUploadZakatOpt Then
        
        
        
        
        s = " SELECT COUNT(PropertyDueBatchDetail.id) AS noofinvoices "
        s = s & " FROM dbo.PropertyDueBatchDetail "
         
            s = s & " INNER JOIN PropertyContractBatch ON PropertyContractBatch.Id = PropertyDueBatchDetail.PropertyContractBatchId "
            s = s & " INNER JOIN PropertyContract ON PropertyContract.Id = PropertyContractBatch.MainDocId "
        s = s & " WHERE 1 = 1 "
        s = s & " AND ISNULL(PropertyDueBatchDetail.zatcaStatus, 0) = 1 "
        
        s = s & " AND ISNULL(IsSelected, 0) = 1"
       ' s = s & "     and PropertyContract.id not in (Select PropertyContractTermination.PropertyContractId from PropertyContractTermination)"
            If Not IsNull(FromDate.value) Then
                s = s & " AND PropertyContractBatch.BatchDate >= " & SQLDate(FromDate, True) & " "
            End If
            
            If Not IsNull(ToDate.value) Then
                s = s & " AND PropertyContractBatch.BatchDate <= " & SQLDate(ToDate, True) & " "
            End If
            
            ' ›· — Õ«·… ZATCA
            s = s & " AND ISNULL(PropertyDueBatchDetail.zatcaStatus, 0) = 1 "
        
        Set rsDummy = New ADODB.Recordset
        
            
                  '
        rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
        If rsDummy.RecordCount > 0 Then
            mCount11 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
        End If
        
        rsDummy.Close
    End If

Label20.Caption = mCount3 + mCount4 + mCount6 + mCount7 + mCount8 + mCount9 + mCount10 + mCount11


   
   
    
       s = " SELECT count(Transaction_ID) as noofinvoices  "
      s = s & "          From dbo.transactions"




  s = s & "  Where (dbo.transactions.Transaction_Type = 21 or dbo.transactions.Transaction_Type = 9)"

  s = s & "   and isnull( Transactions.zatcaStatus,0)=0   "

    If SystemOptions.ZacatHandW Then
        s = s & " and dbo.Transactions.Transaction_Type=854798 "
    End If
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & " "
    End If

    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & " "
    End If
       
    If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     Transactions.BranchId In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
    
  
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and     Transactions.BranchId = " & val(DcBranches(3).BoundText)
        
        
    End If
         
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount3 = IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value)

    End If
    rsDummy.Close

    
          s = " SELECT count(ID) as noofinvoices  "
      s = s & "          From dbo.project_billl"
 

     
 
  s = s & "  Where 1 = 1 and project_billl.bill_type = 0"
                        
    s = s & "   and isnull( project_billl.zatcaStatus,0)=0   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.project_billl.bill_date >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.project_billl.bill_date <=" & SQLDate(ToDate, True) & " "
    End If
    
       If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     project_billl.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
        
     
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and     project_billl.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If
        
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
        mCount4 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    End If
    rsDummy.Close
    
    
    
    
             s = " SELECT count(NoteID) as noofinvoices  "
      s = s & "          From dbo.notes_all"
 

     
 
  s = s & "  Where (dbo.notes_all.NoteType = 85 or dbo.notes_all.NoteType =85)"
s = s & "   and isnull( notes_all.zatcaStatus,0)=0   "
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.notes_all.NoteDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.notes_all.NoteDate  <=" & SQLDate(ToDate, True) & " "
    End If
    
                If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     notes_all.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
       
    
            
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       notes_all.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount8 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
    
           s = " SELECT count(NoteID) as noofinvoices  "
      s = s & "          From dbo.Notes"
 

     
 
  s = s & "  Where (dbo.Notes.NoteType = 9083 or dbo.Notes.NoteType =9082)"
s = s & "   and isnull( Notes.zatcaStatus,0)=0   "
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.Notes.NoteDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.Notes.NoteDate  <=" & SQLDate(ToDate, True) & " "
    End If
    
                If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     notes.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
       
    
            
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       notes.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount6 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
       
           s = " SELECT count(InvoiceID) as noofinvoices  "
      s = s & "          From dbo.tblEInvoice"
 

     
 
  s = s & "  Where 1 = 1"
s = s & "   and isnull( tblEInvoice.zatcaStatus,0)=0   "
  If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     tblEInvoice.branch_id In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
  
  
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       tblEInvoice.branch_id = " & val(DcBranches(3).BoundText)
        
        
    End If
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate  <=" & SQLDate(ToDate, True) & " "
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount7 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
    
     
    
          s = " SELECT count(ID) as noofinvoices  "
      s = s & "          From dbo.TblHandWages"
 

     
 
  s = s & "  Where 1 = 1 "
s = s & "   and isnull( TblHandWages.zatcaStatus,0)=1   "
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.TblHandWages.recordDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.TblHandWages.recordDate  <=" & SQLDate(ToDate, True) & " "
    End If
    s = s & "   and isnull( TblHandWages.zatcaStatus,0)=0   "
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount9 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
    
    
  

s = " SELECT COUNT(ID) AS noofinvoices "
s = s & " FROM dbo.tblContractInsAllocationsDetails "

s = s & " WHERE 1 = 1 "
s = s & " AND ISNULL(tblContractInsAllocationsDetails.zatcaStatus, 0) = 0 "

If Not IsNull(FromDate.value) Then
    s = s & " AND tblContractInsAllocationsDetails.Installdate >= " & SQLDate(FromDate, True) & " "
End If

If Not IsNull(ToDate.value) Then
    s = s & " AND tblContractInsAllocationsDetails.Installdate <= " & SQLDate(ToDate, True) & " "
End If

Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly

If rsDummy.RecordCount > 0 Then
    mCount11 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
End If

rsDummy.Close

Dim mCount12 As Integer

 If SystemOptions.CanUploadZakatOpt Then
        
        
        
        
        s = " SELECT COUNT(PropertyDueBatchDetail.id) AS noofinvoices "
        s = s & " FROM dbo.PropertyDueBatchDetail "
         
            s = s & " INNER JOIN PropertyContractBatch ON PropertyContractBatch.Id = PropertyDueBatchDetail.PropertyContractBatchId "
            s = s & " INNER JOIN PropertyContract ON PropertyContract.Id = PropertyContractBatch.MainDocId "
        s = s & " WHERE 1 = 1 "
        s = s & " AND ISNULL(PropertyDueBatchDetail.zatcaStatus, 0) = 0 "
        s = s & " AND ISNULL(IsSelected, 0) = 1"
's = s & "     and PropertyContract.id not in (Select PropertyContractTermination.PropertyContractId from PropertyContractTermination)"
            If Not IsNull(FromDate.value) Then
                s = s & " AND PropertyContractBatch.BatchDate >= " & SQLDate(FromDate, True) & " "
            End If
            
            If Not IsNull(ToDate.value) Then
                s = s & " AND PropertyContractBatch.BatchDate <= " & SQLDate(ToDate, True) & " "
            End If
            
            ' ›· — Õ«·… ZATCA
            s = s & " AND ISNULL(PropertyDueBatchDetail.zatcaStatus, 0) = 0 "
        
        Set rsDummy = New ADODB.Recordset
        
            
                  '
        rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
        If rsDummy.RecordCount > 0 Then
            mCount12 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
        End If
        
        rsDummy.Close
    End If


    Label22.Caption = mCount3 + mCount4 + mCount6 + mCount8 + mCount7 + mCount9 + mCount11 + mCount12
    
    
    
    
       s = " SELECT count(Transaction_ID) as noofinvoices  "
      s = s & "          From dbo.transactions"
 

     
 
  s = s & "  Where (dbo.transactions.Transaction_Type = 21 or dbo.transactions.Transaction_Type = 9)"
                        
  s = s & "   and isnull( Transactions.zatcaStatus,0)=1   and warrningmessage<>''   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & " "
    End If
       
    If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     Transactions.BranchId In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
  
       
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and      Transactions.BranchId= " & val(DcBranches(3).BoundText)
        
        
    End If
        
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
     mCount3 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    End If
    rsDummy.Close
    
    
    
        s = " SELECT count(ID) as noofinvoices  "
      s = s & "          From dbo.project_billl"
 

     
 
  s = s & "  Where 1 = 1 and project_billl.bill_type = 0"
                        
    s = s & "   and isnull( project_billl.zatcaStatus,0)=1   and warrningmessage<>''   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.project_billl.bill_date >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.project_billl.bill_date <=" & SQLDate(ToDate, True) & " "
    End If
       If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     project_billl.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and      project_billl.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If
        
        
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
        mCount4 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    End If
    rsDummy.Close
    
    
    
       
           s = " SELECT count(NoteID) as noofinvoices  "
      s = s & "          From dbo.Notes"
 

     
 
  s = s & "  Where (dbo.Notes.NoteType = 9083 or dbo.Notes.NoteType =9082)"
s = s & "   and isnull( Notes.zatcaStatus,0)=1   and warrningmessage<>''   "
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.Notes.NoteDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.Notes.NoteDate  <=" & SQLDate(ToDate, True) & " "
    End If
    
                If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     notes.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
       
    
            
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       notes.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount6 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
    
         s = " SELECT count(NoteID) as noofinvoices  "
      s = s & "          From dbo.notes_all"
 

     
 
  s = s & "  Where (dbo.notes_all.NoteType = 85 or dbo.notes_all.NoteType =85)"
s = s & "   and isnull( notes_all.zatcaStatus,0)=1   and warrningmessage<>''   "
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.notes_all.NoteDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.notes_all.NoteDate  <=" & SQLDate(ToDate, True) & " "
    End If
                If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     notes_all.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
       
    
            
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       notes_all.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount8 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    'tblEInvoice.IssueDate
    
         s = " SELECT count(InvoiceID) as noofinvoices  "
      s = s & "          From dbo.tblEInvoice"
 

     
 
  s = s & "  Where 1 = 1"
s = s & "   and isnull( tblEInvoice.zatcaStatus,0)=1   and warrningmessage<>''   "
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate  <=" & SQLDate(ToDate, True) & " "
    End If
      If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     tblEInvoice.branch_id In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
  
  
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       tblEInvoice.branch_id = " & val(DcBranches(3).BoundText)
        
        
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    mCount7 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
    
    
       s = " SELECT count(ID) as noofinvoices  "
      s = s & "          From dbo.TblHandWages"
 

     
 
  s = s & "  Where 1 = 1"
s = s & "   and isnull( TblHandWages.zatcaStatus,0)=1   and warrningmessage<>''   "
' s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.TblHandWages.recordDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.TblHandWages.recordDate  <=" & SQLDate(ToDate, True) & " "
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
         Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
    
    
    
    
    
    mCount9 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
    'Label18.Caption =
    End If
    rsDummy.Close
    
    
    
   'Dim mCount12 As Double

s = " SELECT COUNT(ID) AS noofinvoices "
s = s & " FROM dbo.tblContractInsAllocationsDetails "

s = s & " WHERE 1 = 1 "
s = s & " AND ISNULL(tblContractInsAllocationsDetails.zatcaStatus, 0) = 1 AND warrningmessage <> '' "

If Not IsNull(FromDate.value) Then
    s = s & " AND tblContractInsAllocationsDetails.Installdate >= " & SQLDate(FromDate, True) & " "
End If

If Not IsNull(ToDate.value) Then
    s = s & " AND tblContractInsAllocationsDetails.Installdate <= " & SQLDate(ToDate, True) & " "
End If

Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly

If rsDummy.RecordCount > 0 Then
    mCount12 = val(IIf(IsNull(rsDummy("noofinvoices").value), "", rsDummy("noofinvoices").value))
End If

rsDummy.Close
 
    
    Label26.Caption = mCount3 + mCount4 + mCount6 + mCount8 + mCount7 + mCount9 + mCount12
     
End Function


Public Function GetResults() As Boolean
    On Error GoTo ErrTrap

    Dim s As String
    Dim rsDummy As ADODB.Recordset
    Set rsDummy = New ADODB.Recordset

    '================= Counters =================
    Dim mCount1 As Double, mCount2 As Double, mCount3 As Double, mCount4 As Double
    Dim mCount5 As Double, mCount6 As Double, mCount7 As Double, mCount8 As Double
    Dim mCount9 As Double, mCount10 As Double, mCount11 As Double, mCount12 As Double
    Dim mCount111 As Double

    ' « ’«· “ﬂ« « ≈‰ ·“„
    If SystemOptions.CanUploadZakatOpt Then
        ConectionFirst
    End If

    '------------------------[A] ≈Ã„«·Ì »œÊ‰ ‘—ÿ «·Õ«·… ------------------------
    ' Transactions (›« Ê—… »Ì⁄/‘—«¡ ﬂ«‘)
    s = "SELECT COUNT(Transaction_ID) AS noofinvoices FROM dbo.Transactions WHERE (Transaction_Type = 21 )"
    If SystemOptions.ZacatHandW Then
        s = s & " AND Transactions.Transaction_Type = 854798"
    End If
    If Not IsNull(FromDate.value) Then s = s & " AND Transactions.Transaction_Date >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND Transactions.Transaction_Date <= " & SQLDate(ToDate, True)

    If Trim(DcBranches(2).Text) <> "" Then
        s = s & " AND Transactions.BranchId IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    End If
    If Trim(DcBranches(3).Text) <> "" Then
        s = s & " AND Transactions.BranchId = " & val(DcBranches(3).BoundText)
    End If

    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then mCount1 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
    rsDummy.Close

    ' Notes (9082/9083)
    s = "SELECT COUNT(NoteID) AS noofinvoices FROM dbo.Notes WHERE (NoteType = 9083 OR NoteType = 9082)"
    If Not IsNull(FromDate.value) Then s = s & " AND NoteDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND NoteDate <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then
        s = s & " AND Notes.branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    End If
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND Notes.branch_no = " & val(DcBranches(3).BoundText)

    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then mCount5 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
    rsDummy.Close

    ' notes_all (85)
    s = "SELECT COUNT(NoteID) AS noofinvoices FROM dbo.notes_all WHERE (NoteType = 85)"
    If Not IsNull(FromDate.value) Then s = s & " AND NoteDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND NoteDate <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then
        s = s & " AND notes_all.branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    End If
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND notes_all.branch_no = " & val(DcBranches(3).BoundText)

    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then mCount7 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
    rsDummy.Close

    ' project_billl (›Ê« Ì— «·„‘«—Ì⁄)
    s = "SELECT COUNT(ID) AS noofinvoices FROM dbo.project_billl WHERE bill_type = 0"
    If Not IsNull(FromDate.value) Then s = s & " AND bill_date >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND bill_date <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then
        s = s & " AND project_billl.branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    End If
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND project_billl.branch_no = " & val(DcBranches(3).BoundText)

    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then mCount2 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
    rsDummy.Close

    ' tblEInvoice («·≈·ﬂ —Ê‰Ì…)
    s = "SELECT COUNT(InvoiceID) AS noofinvoices FROM dbo.tblEInvoice WHERE 1=1"
    If Not IsNull(FromDate.value) Then s = s & " AND IssueDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND IssueDate <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then
        s = s & " AND tblEInvoice.branch_id IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    End If
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND tblEInvoice.branch_id = " & val(DcBranches(3).BoundText)

    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then mCount8 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
    rsDummy.Close

    ' PropertyDueBatchDetail (ZATCA ñ „Œ «—… ›ﬁÿ) ⁄»— POSConnection
    mCount111 = 0
    Dim mWebAll As Double
    If SystemOptions.CanUploadZakatOpt Then
        s = ""
        s = s & "SELECT COUNT(pdbd.ID) AS noofinvoices" & vbCrLf
        s = s & "FROM dbo.PropertyDueBatchDetail AS pdbd" & vbCrLf
        s = s & "INNER JOIN dbo.PropertyContractBatch AS pcb ON pcb.Id = pdbd.PropertyContractBatchId" & vbCrLf
        s = s & "INNER JOIN dbo.PropertyContract AS pc ON pc.Id = pcb.MainDocId" & vbCrLf
        s = s & "WHERE ISNULL(pdbd.IsSelected,0) = 1" & vbCrLf
        If Not IsNull(FromDate.value) Then s = s & "AND pcb.BatchDate >= " & SQLDate(FromDate, True) & vbCrLf
        If Not IsNull(ToDate.value) Then s = s & "AND pcb.BatchDate <= " & SQLDate(ToDate, True) & vbCrLf

        rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
        If rsDummy.RecordCount > 0 Then mCount111 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
        rsDummy.Close
        
        
                
                ' ===== DebitAndCreditNotification: ≈Ã„«·Ì »œÊ‰ Õ«·… =====
        
        s = "SELECT COUNT(d.Id) AS noofinvoices FROM dbo.DebitAndCreditNotification d WHERE 1=1"
        If Not IsNull(FromDate.value) Then s = s & " AND d.[Date] >= " & SQLDate(FromDate, True)
        If Not IsNull(ToDate.value) Then s = s & " AND d.[Date] <= " & SQLDate(ToDate, True)
        
        ' ›· — «·‰‘«ÿ (ActivityTypeId) ⁄·Ï ÃœÊ· Department
        If Trim(DcBranches(2).Text) <> "" Then
            s = s & " AND d.DepartmentId IN (SELECT Id FROM dbo.Department WHERE ActivityId = " & val(DcBranches(2).BoundText) & ")"
        End If
        ' ›· — «·›—⁄ «·„»«‘—
        If Trim(DcBranches(3).Text) <> "" Then
            s = s & " AND d.DepartmentId = " & val(DcBranches(3).BoundText)
        End If
        
        rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
        If rsDummy.RecordCount > 0 Then mWebAll = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
        rsDummy.Close

    End If



' ===== DebitAndCreditNotification: ?????? ???? ???? =====

    Label18.Caption = mCount1 + mCount2 + mCount5 + mCount7 + mCount8 + mCount111 + mWebAll

    '------------------------[B] zatcaStatus = 1 („ı⁄ „œ) ------------------------
    mCount3 = 0: mCount4 = 0: mCount6 = 0: mCount7 = 0: mCount8 = 0: mCount9 = 0: mCount10 = 0: mCount11 = 0

    ' Transactions
    s = "SELECT COUNT(Transaction_ID) AS noofinvoices FROM dbo.Transactions WHERE (Transaction_Type = 21 ) AND ISNULL(zatcaStatus,0)=1"
    If SystemOptions.ZacatHandW Then s = s & " AND Transactions.Transaction_Type = 854798"
    If Not IsNull(FromDate.value) Then s = s & " AND Transaction_Date >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND Transaction_Date <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND BranchId IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND BranchId = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount3 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' project_billl
    s = "SELECT COUNT(ID) AS noofinvoices FROM dbo.project_billl WHERE bill_type = 0 AND ISNULL(zatcaStatus,0)=1"
    If Not IsNull(FromDate.value) Then s = s & " AND bill_date >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND bill_date <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_no = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount4 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' Notes (9082/9083)
    s = "SELECT COUNT(NoteID) AS noofinvoices FROM dbo.Notes WHERE (NoteType IN (9082,9083)) AND ISNULL(zatcaStatus,0)=1"
    If Not IsNull(FromDate.value) Then s = s & " AND NoteDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND NoteDate <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_no = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount6 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' notes_all (85)
    s = "SELECT COUNT(NoteID) AS noofinvoices FROM dbo.notes_all WHERE NoteType = 85 AND ISNULL(zatcaStatus,0)=1"
    If Not IsNull(FromDate.value) Then s = s & " AND NoteDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND NoteDate <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_no = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount7 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' tblEInvoice
    s = "SELECT COUNT(InvoiceID) AS noofinvoices FROM dbo.tblEInvoice WHERE ISNULL(zatcaStatus,0)=1"
    If Not IsNull(FromDate.value) Then s = s & " AND IssueDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND IssueDate <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_id IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_id = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount8 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' TblHandWages
    s = "SELECT COUNT(ID) AS noofinvoices FROM dbo.TblHandWages WHERE ISNULL(zatcaStatus,0)=1"
    If Not IsNull(FromDate.value) Then s = s & " AND recordDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND recordDate <= " & SQLDate(ToDate, True)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount9 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' tblContractInsAllocationsDetails
    s = "SELECT COUNT(ID) AS noofinvoices FROM dbo.tblContractInsAllocationsDetails WHERE ISNULL(zatcaStatus,0)=1"
    If Not IsNull(FromDate.value) Then s = s & " AND Installdate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND Installdate <= " & SQLDate(ToDate, True)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount10 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' PropertyDueBatchDetail ⁄»— POSConnection
    Dim mWebApproved As Double
    If SystemOptions.CanUploadZakatOpt Then
        s = ""
        s = s & "SELECT COUNT(pdbd.ID) AS noofinvoices" & vbCrLf
        s = s & "FROM dbo.PropertyDueBatchDetail AS pdbd" & vbCrLf
        s = s & "INNER JOIN dbo.PropertyContractBatch AS pcb ON pcb.Id = pdbd.PropertyContractBatchId" & vbCrLf
        s = s & "INNER JOIN dbo.PropertyContract AS pc ON pc.Id = pcb.MainDocId" & vbCrLf
        s = s & "WHERE ISNULL(pdbd.IsSelected,0)=1 AND ISNULL(pdbd.zatcaStatus,0)=1" & vbCrLf
        If Not IsNull(FromDate.value) Then s = s & "AND pcb.BatchDate >= " & SQLDate(FromDate, True) & vbCrLf
        If Not IsNull(ToDate.value) Then s = s & "AND pcb.BatchDate <= " & SQLDate(ToDate, True) & vbCrLf

        rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
        If rsDummy.RecordCount > 0 Then mCount11 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
        rsDummy.Close
        
        ' ===== DebitAndCreditNotification: „⁄ „œ (zatcaStatus=1) =====
        
        s = "SELECT COUNT(d.Id) AS noofinvoices FROM dbo.DebitAndCreditNotification d WHERE ISNULL(d.zatcaStatus,0)=1"
        If Not IsNull(FromDate.value) Then s = s & " AND d.[Date] >= " & SQLDate(FromDate, True)
        If Not IsNull(ToDate.value) Then s = s & " AND d.[Date] <= " & SQLDate(ToDate, True)
        If Trim(DcBranches(2).Text) <> "" Then
            s = s & " AND d.DepartmentId IN (SELECT Id FROM dbo.Department WHERE ActivityId = " & val(DcBranches(2).BoundText) & ")"
        End If
        If Trim(DcBranches(3).Text) <> "" Then
            s = s & " AND d.DepartmentId = " & val(DcBranches(3).BoundText)
        End If
        rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
        If rsDummy.RecordCount > 0 Then mWebApproved = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
        rsDummy.Close

    End If

    Label20.Caption = mCount3 + mCount4 + mCount6 + mCount7 + mCount8 + mCount9 + mCount10 + mCount11 + mWebApproved

    '------------------------[C] zatcaStatus = 0 (€Ì— „ı⁄ „œ) ------------------------
    mCount3 = 0: mCount4 = 0: mCount6 = 0: mCount7 = 0: mCount8 = 0: mCount9 = 0: mCount11 = 0: mCount12 = 0

    ' Transactions
    s = "SELECT COUNT(Transaction_ID) AS noofinvoices FROM dbo.Transactions WHERE (Transaction_Type = 21 ) AND ISNULL(zatcaStatus,0)=0"
    If SystemOptions.ZacatHandW Then s = s & " AND Transactions.Transaction_Type = 854798"
    If Not IsNull(FromDate.value) Then s = s & " AND Transaction_Date >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND Transaction_Date <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND BranchId IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND BranchId = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount3 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' project_billl
    s = "SELECT COUNT(ID) AS noofinvoices FROM dbo.project_billl WHERE bill_type = 0 AND ISNULL(zatcaStatus,0)=0"
    If Not IsNull(FromDate.value) Then s = s & " AND bill_date >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND bill_date <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_no = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount4 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' Notes
    s = "SELECT COUNT(NoteID) AS noofinvoices FROM dbo.Notes WHERE (NoteType IN (9082,9083)) AND ISNULL(zatcaStatus,0)=0"
    If Not IsNull(FromDate.value) Then s = s & " AND NoteDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND NoteDate <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_no = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount6 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' notes_all
    s = "SELECT COUNT(NoteID) AS noofinvoices FROM dbo.notes_all WHERE NoteType = 85 AND ISNULL(zatcaStatus,0)=0"
    If Not IsNull(FromDate.value) Then s = s & " AND NoteDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND NoteDate <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_no = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount8 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' tblEInvoice
    s = "SELECT COUNT(InvoiceID) AS noofinvoices FROM dbo.tblEInvoice WHERE ISNULL(zatcaStatus,0)=0"
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_id IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_id = " & val(DcBranches(3).BoundText)
    If Not IsNull(FromDate.value) Then s = s & " AND IssueDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND IssueDate <= " & SQLDate(ToDate, True)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount7 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' TblHandWages
    s = "SELECT COUNT(ID) AS noofinvoices FROM dbo.TblHandWages WHERE ISNULL(zatcaStatus,0)=0"
    If Not IsNull(FromDate.value) Then s = s & " AND recordDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND recordDate <= " & SQLDate(ToDate, True)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount9 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' tblContractInsAllocationsDetails
    s = "SELECT COUNT(ID) AS noofinvoices FROM dbo.tblContractInsAllocationsDetails WHERE ISNULL(zatcaStatus,0)=0"
    If Not IsNull(FromDate.value) Then s = s & " AND Installdate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND Installdate <= " & SQLDate(ToDate, True)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount11 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' PropertyDueBatchDetail ⁄»— POSConnection
    Dim mWebPending As Double
    If SystemOptions.CanUploadZakatOpt Then
        s = ""
        s = s & "SELECT COUNT(pdbd.ID) AS noofinvoices" & vbCrLf
        s = s & "FROM dbo.PropertyDueBatchDetail AS pdbd" & vbCrLf
        s = s & "INNER JOIN dbo.PropertyContractBatch AS pcb ON pcb.Id = pdbd.PropertyContractBatchId" & vbCrLf
        s = s & "INNER JOIN dbo.PropertyContract AS pc ON pc.Id = pcb.MainDocId" & vbCrLf
        s = s & "WHERE ISNULL(pdbd.IsSelected,0)=1 AND ISNULL(pdbd.zatcaStatus,0)=0" & vbCrLf
        If Not IsNull(FromDate.value) Then s = s & "AND pcb.BatchDate >= " & SQLDate(FromDate, True) & vbCrLf
        If Not IsNull(ToDate.value) Then s = s & "AND pcb.BatchDate <= " & SQLDate(ToDate, True) & vbCrLf

        rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
        If rsDummy.RecordCount > 0 Then mCount12 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
        rsDummy.Close
        
        
        ' ===== DebitAndCreditNotification: €Ì— „⁄ „œ (zatcaStatus=0) =====
        
        s = "SELECT COUNT(d.Id) AS noofinvoices FROM dbo.DebitAndCreditNotification d WHERE ISNULL(d.zatcaStatus,0)=0"
        If Not IsNull(FromDate.value) Then s = s & " AND d.[Date] >= " & SQLDate(FromDate, True)
        If Not IsNull(ToDate.value) Then s = s & " AND d.[Date] <= " & SQLDate(ToDate, True)
        If Trim(DcBranches(2).Text) <> "" Then
            s = s & " AND d.DepartmentId IN (SELECT Id FROM dbo.Department WHERE ActivityId = " & val(DcBranches(2).BoundText) & ")"
        End If
        If Trim(DcBranches(3).Text) <> "" Then
            s = s & " AND d.DepartmentId = " & val(DcBranches(3).BoundText)
        End If
        rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
        If rsDummy.RecordCount > 0 Then mWebPending = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
        rsDummy.Close

    End If

    Label22.Caption = mCount3 + mCount4 + mCount6 + mCount8 + mCount7 + mCount9 + mCount11 + mCount12 + mWebPending

    '------------------------[D] «·„ﬁ»Ê· „⁄  Õ–Ì— (warningmessage <> '') ------------------------
    mCount3 = 0: mCount4 = 0: mCount6 = 0: mCount7 = 0: mCount9 = 0: mCount12 = 0

    ' Transactions
    s = "SELECT COUNT(Transaction_ID) AS noofinvoices FROM dbo.Transactions WHERE (Transaction_Type IN (21,9)) AND ISNULL(zatcaStatus,0)=1 AND ISNULL(warrningmessage,'')<>''"
    If Not IsNull(FromDate.value) Then s = s & " AND Transaction_Date >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND Transaction_Date <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND BranchId IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND BranchId = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount3 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' project_billl
    s = "SELECT COUNT(ID) AS noofinvoices FROM dbo.project_billl WHERE bill_type = 0 AND ISNULL(zatcaStatus,0)=1 AND ISNULL(warrningmessage,'')<>''"
    If Not IsNull(FromDate.value) Then s = s & " AND bill_date >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND bill_date <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_no = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount4 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' Notes
    s = "SELECT COUNT(NoteID) AS noofinvoices FROM dbo.Notes WHERE (NoteType IN (9082,9083)) AND ISNULL(zatcaStatus,0)=1 AND ISNULL(warrningmessage,'')<>''"
    If Not IsNull(FromDate.value) Then s = s & " AND NoteDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND NoteDate <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_no = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount6 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' notes_all
    s = "SELECT COUNT(NoteID) AS noofinvoices FROM dbo.notes_all WHERE NoteType = 85 AND ISNULL(zatcaStatus,0)=1 AND ISNULL(warrningmessage,'')<>''"
    If Not IsNull(FromDate.value) Then s = s & " AND NoteDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND NoteDate <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_no IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_no = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount7 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' tblEInvoice
    s = "SELECT COUNT(InvoiceID) AS noofinvoices FROM dbo.tblEInvoice WHERE ISNULL(zatcaStatus,0)=1 AND ISNULL(warrningmessage,'')<>''"
    If Not IsNull(FromDate.value) Then s = s & " AND IssueDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND IssueDate <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then s = s & " AND branch_id IN (SELECT branch_id FROM dbo.TblBranchesData WHERE ActivityTypeId = " & val(DcBranches(2).BoundText) & ")"
    If Trim(DcBranches(3).Text) <> "" Then s = s & " AND branch_id = " & val(DcBranches(3).BoundText)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount8 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' TblHandWages
    s = "SELECT COUNT(ID) AS noofinvoices FROM dbo.TblHandWages WHERE ISNULL(zatcaStatus,0)=1 AND ISNULL(warrningmessage,'')<>''"
    If Not IsNull(FromDate.value) Then s = s & " AND recordDate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND recordDate <= " & SQLDate(ToDate, True)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount9 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close

    ' tblContractInsAllocationsDetails (status=1 „⁄  Õ–Ì—)
    s = "SELECT COUNT(ID) AS noofinvoices FROM dbo.tblContractInsAllocationsDetails WHERE ISNULL(zatcaStatus,0)=1 AND ISNULL(warrningmessage,'')<>''"
    If Not IsNull(FromDate.value) Then s = s & " AND Installdate >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND Installdate <= " & SQLDate(ToDate, True)
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly: If rsDummy.RecordCount > 0 Then mCount12 = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value)): rsDummy.Close





' ===== PropertyDueBatchDetail: „⁄ „œ Ê»œ«Œ·Â  Õ–Ì— (POS/ZATCA) =====
Dim mPropWarn As Double
Dim mWebWarn As Double
mPropWarn = 0

If SystemOptions.CanUploadZakatOpt Then
    s = "SELECT COUNT(p.Id) AS noofinvoices " & _
        "FROM dbo.PropertyDueBatchDetail p " & _
        "INNER JOIN dbo.PropertyContractBatch pcb ON pcb.Id = p.PropertyContractBatchId " & _
        "INNER JOIN dbo.PropertyContract pc ON pc.Id = pcb.MainDocId " & _
        "LEFT JOIN dbo.Department dp ON dp.Id = pc.DepartmentId " & _
        "WHERE 1=1 " & _
        "AND ISNULL(p.zatcaStatus,0)=1 " & _
        "AND ISNULL(p.warrningmessage,'')<>'' " & _
        "AND ISNULL(p.IsSelected,0)=1"

    If Not IsNull(FromDate.value) Then
        s = s & " AND pcb.BatchDate >= " & SQLDate(FromDate, True)
    End If
    If Not IsNull(ToDate.value) Then
        s = s & " AND pcb.BatchDate <= " & SQLDate(ToDate, True)
    End If

    ' ›· — «·‰‘«ÿ (ActivityId) „‰ Department
    If Trim(DcBranches(2).Text) <> "" Then
        s = s & " AND dp.ActivityId = " & val(DcBranches(2).BoundText)
    End If
    ' ›· — «·›—⁄ (DepartmentId)
    If Trim(DcBranches(3).Text) <> "" Then
        s = s & " AND dp.Id = " & val(DcBranches(3).BoundText)
    End If

    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then
        mPropWarn = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
    End If
    rsDummy.Close
    
    
    ' ===== DebitAndCreditNotification: „⁄ „œ Ê»œ«Œ·Â  Õ–Ì— =====
    
    s = "SELECT COUNT(d.Id) AS noofinvoices FROM dbo.DebitAndCreditNotification d WHERE ISNULL(d.zatcaStatus,0)=1 AND ISNULL(d.warrningmessage,'')<>''"
    If Not IsNull(FromDate.value) Then s = s & " AND d.[Date] >= " & SQLDate(FromDate, True)
    If Not IsNull(ToDate.value) Then s = s & " AND d.[Date] <= " & SQLDate(ToDate, True)
    If Trim(DcBranches(2).Text) <> "" Then
        s = s & " AND d.DepartmentId IN (SELECT Id FROM dbo.Department WHERE ActivityId = " & val(DcBranches(2).BoundText) & ")"
    End If
    If Trim(DcBranches(3).Text) <> "" Then
        s = s & " AND d.DepartmentId = " & val(DcBranches(3).BoundText)
    End If
    rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
    If rsDummy.RecordCount > 0 Then mWebWarn = val(IIf(IsNull(rsDummy("noofinvoices").value), 0, rsDummy("noofinvoices").value))
    rsDummy.Close

End If


    Label26.Caption = mCount3 + mCount4 + mCount6 + mCount8 + mCount7 + mCount9 + mCount12 + mPropWarn + mWebWarn

    GetResults = True
    Set rsDummy = Nothing
    Exit Function

ErrTrap:
    GetResults = False
    On Error Resume Next
    If Not rsDummy Is Nothing Then
        If rsDummy.State <> 0 Then rsDummy.Close
        Set rsDummy = Nothing
    End If
End Function


Private Sub CmdDelete_Click()
 
Dim i As Long, IntCounter As Long
Dim s As String
 With Me.grd(0)
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked And grd(0).cell(flexcpBackColor, i, 1, i, 56) = vbRed Then
                    s = "delete from  tblEInvoice   where InvoiceID=" & val(.TextMatrix(i, .ColIndex("ID")))
                    Cn.Execute s
                    s = "delete from  tblEInvoice2   where InvoiceID=" & val(.TextMatrix(i, .ColIndex("ID")))
                    Cn.Execute s
                    
                End If
        Next
End With
cmdInsert_Click
End Sub

Private Sub cmdInsert_Click()
GetResults
               
    
    
      'PaymentMeansCode
            '10 In cash
            '30 Credit
            '42 Payment to bank account
            '48 Bank card
            '1 Instrument not defined(Free text)
               'for information
            ' VAT Category (O) "Not subject to VAT" (O)  €Ì— Œ«÷⁄ ·÷—Ì»… «·ﬁÌ„… «·„÷«›… ·«»œ «‰ ÌﬂÊ‰ ‰”»… «·÷—Ì»… ’›—
            ' VAT Category (E)   „⁄›Ï „‰ ÷—Ì»… «·ﬁÌ„… «·„÷«›… ‰”»… «·÷—Ì»… Â ﬂÊ‰ ’›— Ê·«»œ „‰ –ﬂ— ”»» «·«⁄›«¡ :TaxExemptionReason
            ' VAT Category (S)   ÷—Ì»… «·ﬁÌ„… «·„÷«›… ·«»œ „‰ ﬂ «»… «·‰”»… Ê ﬂÊ‰ «ﬂ»— „‰ ’›—
            ' VAT Category (Z)   Zero rated goods
            
            
              ' ›« Ê—… ÷—Ì»Ì… «Ê „»”ÿ… 388
      ' «‘⁄«— „œÌ‰ 383
      ' 381 «‘⁄«— œ«∆‰
     
    'inv.invoiceTypeCode.Name based on format NNPNESB
    'NN 01 ··›« Ê—… «·÷—Ì»Ì…
    'NN 02 ··›« Ê—… «·÷—Ì»Ì… «·„»”ÿ…
    'P ›Ï Õ«·… ›« Ê—… ·ÿ—› À«·À ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'N ›Ï Õ«·… ›« Ê—… «”„Ì… ‰ﬂ » 1 ›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'E ›Ï Õ«·… ›« Ê—… ··’«œ—«  ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'S ›Ï Õ«·… ›« Ê—… „·Œ’… ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'B  ›Ï Õ«·… ›« Ê—… –« Ì… ‰ﬂ » 1
    'B ›Ï Õ«·… «‰ «·›« Ê—… ’«œ—« =1 ·« Ì„ﬂ‰ «‰  ﬂÊ‰ «·›« Ê—… –« Ì… =1
     
     
Dim s As String
Dim rsDummy As New ADODB.Recordset

    s = " SELECT   tblActivitesType.id as ActivityTypeId, 'Sales Invoice' as TypeName, TblBranchesData.branch_id,TblBranchesData.branch_name BranchName, tblActivitesType.Name as ActivityName, Transactions.Invoicetype,dbo.Transactions.ErrorMessageS , Transactions.chkTaxExempt, dbo.Transactions.DateBaptizing,dbo.Transactions.NoteSerial1 order_no,  dbo.Transactions.ReturnSerial,  dbo.Transactions.SalesInvoiceDate ,  dbo.Transactions.Transaction_ID, PayeeFinancialAccount =(select IBan  from BanksData  where bankid=Transactions.bankid),     dbo.Transactions.Transaction_ID AS Expr1, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1 AS id, dbo.Transactions.Transaction_Date AS IssueDate, dbo.Transactions.RecTime AS IssueTim, "
    s = s & "                        dbo.Transactions.InvoiceTypeCodeID, dbo.Transactions.InvoiceTypeCodename, dbo.Transactions.DocumentCurrencyCode, dbo.Transactions.TaxCurrencyCode, dbo.Transactions.InvoiceDocumentReferenceID,"
    s = s & "                          dbo.Transactions.AdditionalDocumentReferenceICVUUID, dbo.Transactions.ActualDeliveryDate, dbo.Transactions.LatestDeliveryDate, dbo.Transactions.PaymentMeansCode, dbo.Transactions.InstructionNote,"
    s = s & "                          dbo.Transactions.paymentnote, dbo.TblCustemers.CustGID AS Identificationid, 'CRN' AS schemeID, dbo.TblCustemers.StreetName, dbo.TblCustemers.AdditionalStreetName, dbo.TblCustemers.BuildingNumber,"
    s = s & "                          dbo.TblCustemers.PlotIdentification, dbo.TblCustemers.CityName, dbo.TblCustemers.PostalZone, dbo.TblCustemers.CountrySubentity, dbo.TblCustemers.CitySubdivisionName, dbo.TblCustemers.IdentificationCode,"
    s = s & "                          dbo.TblCustemers.CusNamee AS RegistrationName, dbo.TblCustemers.VATNO AS CompanyID, dbo.TblCustemers.Id700 ,  dbo.Transactions.LblDiscountsTotal AS allowancechargeAmount, 'Discount' AS AllowanceChargeReason, 'S' AS TaxCategoryID,"
    s = s & "                          '15' AS TaxCategoryPercent, dbo.Transactions.last_changed, dbo.Transactions.Transaction_NetValue AS PayableAmount, dbo.Transactions.AdvPay AS PrepaidAmount, dbo.transactionsVatDetails.SingedXMLFileName,"
    s = s & "                          dbo.transactionsVatDetails.PIH, dbo.transactionsVatDetails.QRCode, dbo.transactionsVatDetails.UUID, dbo.transactionsVatDetails.InvoiceHash, dbo.transactionsVatDetails.EncodedInvoice,"
    s = s & "                          dbo.transactionsVatDetails.SingedXML,  dbo.transactionsVatDetails.QrCodeDataPath,0 as DocType,0 VatValue,TblCustemers.Export"
    s = s & "  FROM            dbo.TblCustemers INNER JOIN"
    s = s & "                          dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID left OUTER JOIN"
    s = s & "                          dbo.transactionsVatDetails ON dbo.Transactions.Transaction_ID = dbo.transactionsVatDetails.Transaction_ID and isnull(transactionsVatDetails.isdeleted,0)=0"
    s = s & " and transactionsVatDetails.TableName = 'Transactions'"
    
    s = s & " Inner join TblBranchesData On TblBranchesData.branch_id = Transactions.BranchId inner join tblActivitesType On tblActivitesType.id =TblBranchesData.ActivityTypeId  "
                         
    s = s & "  Where (dbo.transactions.Transaction_Type = 21 or dbo.transactions.Transaction_Type = 9)"
                        
     s = s & "   and isnull( Transactions.zatcaStatus,0)<>1   "
    If SystemOptions.ZacatHandW Then
        s = s & " and dbo.Transactions.Transaction_Type=854798 "
    End If
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate, True) & " "
    End If
    
    If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     Transactions.BranchId In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
  
  
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       Transactions.BranchId = " & val(DcBranches(3).BoundText)
        
        
    End If
    s = s & " and IsNull(Transactions.IsHiddenVat,0) = 0        "
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)

    
    
     s = s & "  Union all"
    
    s = s & " SELECT tblActivitesType.id as ActivityTypeId,'Credit Or Debit Note' as TypeName,  TblBranchesData.branch_id,TblBranchesData.branch_name BranchName, tblActivitesType.Name as ActivityName, "
    s = s & "     Notes.Invoicetype"
   s = s & " ,dbo.Notes.ErrorMessageS,chkTaxExempt = 0"
   s = s & " ,dbo.Notes.NoteDate DateBaptizing"
   s = s & " ,CAST(dbo.Notes.order_no AS VARCHAR(10)) order_no"
   s = s & " ,order_no ReturnSerial"
   s = s & " ,dbo.Notes.NoteDate SalesInvoiceDate"
   s = s & " ,dbo.Notes.Noteid Transaction_ID"
   s = s & " ,PayeeFinancialAccount = (SELECT"
    s = s & "             IBan"
    s = s & "      From BanksData"
    s = s & "         WHERE bankid = 0)"
   s = s & " ,dbo.Notes.Noteid AS Expr1"
   s = s & " ,Transaction_Type = (CASE Notes.NoteType"
    s = s & "         WHEN 9082 THEN 383"
    s = s & "         WHEN 9083 THEN 381"
    s = s & " END)"

   
   s = s & " ,CAST(dbo.Notes.NoteSerial1 AS NVARCHAR(50)) AS id"
   s = s & " ,dbo.Notes.NoteDate AS IssueDate"
   s = s & " ,dbo.Notes.RecTime AS IssueTim"
   s = s & " ,dbo.Notes.InvoiceTypeCodeID"
   s = s & " ,dbo.Notes.InvoiceTypeCodename"
   s = s & " ,dbo.Notes.DocumentCurrencyCode"
   s = s & " ,dbo.Notes.TaxCurrencyCode"
   s = s & " ,dbo.Notes.InvoiceDocumentReferenceID"
   s = s & " ,dbo.Notes.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.Notes.ActualDeliveryDate"
   s = s & " ,dbo.Notes.LatestDeliveryDate"
   s = s & " ,dbo.Notes.PaymentMeansCode"
   s = s & " ,dbo.Notes.InstructionNote"
   s = s & " ,dbo.Notes.paymentnote"
   s = s & " ,dbo.TblCustemers.CustGID AS Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.TblCustemers.StreetName"
   s = s & " ,dbo.TblCustemers.AdditionalStreetName"
   s = s & " ,dbo.TblCustemers.BuildingNumber"
   s = s & " ,dbo.TblCustemers.PlotIdentification"
   s = s & " ,dbo.TblCustemers.CityName"
   s = s & " ,dbo.TblCustemers.PostalZone"
   s = s & " ,dbo.TblCustemers.CountrySubentity"
   s = s & " ,dbo.TblCustemers.CitySubdivisionName"
   s = s & " ,dbo.TblCustemers.IdentificationCode"
   s = s & " ,dbo.TblCustemers.CusNamee AS RegistrationName"
   s = s & " ,dbo.TblCustemers.VATNO AS CompanyID , dbo.TblCustemers.Id700 "
   s = s & " ,0 AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,'S' AS TaxCategoryID"
   s = s & " ,'15' AS TaxCategoryPercent"
   s = s & " ,Notes.last_changed"
   
   s = s & " ,vat + Note_Value AS PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
   s = s & " ,3 AS DocType,vat as VatValue,TblCustemers.Export"
    s = s & " From dbo.TblCustemers"
s = s & " INNER JOIN dbo.Notes"
s = s & "     ON dbo.TblCustemers.CusID = dbo.notes.CusID"
s = s & " left outer JOIN dbo.transactionsVatDetails"
s = s & " ON dbo.Notes.NoteID = dbo.transactionsVatDetails.Transaction_ID and transactionsVatDetails.TableName = 'Notes'"
s = s & "         AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"
s = s & " Inner join TblBranchesData On TblBranchesData.branch_id = Notes.branch_no inner join tblActivitesType On tblActivitesType.id =TblBranchesData.ActivityTypeId  "

s = s & " Where 1 = 1"
s = s & " AND Notes.NoteType IN (9083, 9082)"

    
  
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.Notes.NoteDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.Notes.NoteDate <=" & SQLDate(ToDate, True) & " "
    End If
    
    s = s & " and ISNULL(Notes.zatcaStatus, 0) <> 1"
   ' s = s & "                      ORDER by  InvoiceTypeCodeID desc ,Transaction_Date, NoteSerial1"
    
    
        
    If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     Notes.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If

      
  
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       Notes.branch_no= " & val(DcBranches(3).BoundText)
    End If
        
    
     s = s & "  Union all"
    
    s = s & " SELECT  '' as ActivityTypeId , 'Sales Invoice Excel' as TypeName,  branch_id,  BranchName, '' as  ActivityName, "
    s = s & "     tblEInvoice.DefaultInvoicetype as Invoicetype"
   s = s & " ,dbo.tblEInvoice.ErrorMessageS,chkTaxExempt "
   s = s & " ,dbo.tblEInvoice.IssueDate DateBaptizing"
   s = s & " ,(dbo.tblEInvoice.InvoiceID ) order_no"
   s = s & " ,'0' ReturnSerial"
   s = s & " ,dbo.tblEInvoice.IssueDate SalesInvoiceDate"
   s = s & " ,dbo.tblEInvoice.ID Transaction_ID"
   s = s & " ,PayeeFinancialAccount = ''"
   s = s & " ,dbo.tblEInvoice.ID AS Expr1"
   s = s & " ,Transaction_Type =21"

   s = s & " ,dbo.tblEInvoice.InvoiceID AS id"
   s = s & " ,dbo.tblEInvoice.IssueDate AS IssueDate"
   s = s & " ,dbo.tblEInvoice.RecTime AS IssueTim"
   s = s & " ,dbo.tblEInvoice.InvoiceTypeCodeID"
   s = s & " ,dbo.tblEInvoice.InvoiceTypeCodename"
   s = s & " ,dbo.tblEInvoice.DocumentCurrencyCode"
   s = s & " ,dbo.tblEInvoice.TaxCurrencyCode"
   s = s & " ,dbo.tblEInvoice.InvoiceDocumentReferenceID"
   s = s & " ,dbo.tblEInvoice.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.tblEInvoice.ActualDeliveryDate"
   s = s & " ,dbo.tblEInvoice.LatestDeliveryDate"
   s = s & " ,dbo.tblEInvoice.PaymentMeansCode"
   s = s & " ,dbo.tblEInvoice.InstructionNote"
   s = s & " ,dbo.tblEInvoice.paymentnote"
   s = s & " ,dbo.tblEInvoice.Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.tblEInvoice.StreetName"
   s = s & " ,dbo.tblEInvoice.AdditionalStreetName"
   s = s & " ,dbo.tblEInvoice.BuildingNumber"
   s = s & " ,dbo.tblEInvoice.PlotIdentification"
   s = s & " ,dbo.tblEInvoice.CityName"
   s = s & " ,dbo.tblEInvoice.PostalZone"
   s = s & " ,dbo.tblEInvoice.CountrySubentity"
   s = s & " ,dbo.tblEInvoice.CitySubdivisionName"
   s = s & " ,dbo.tblEInvoice.IdentificationCode"
   s = s & " ,dbo.tblEInvoice.RegistrationName"
   s = s & " ,dbo.tblEInvoice.CompanyID , dbo.tblEInvoice.Id700 "
   s = s & " ,0 AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,tblEInvoice.TaxCategoryID"
   
   's = s & " ,round(VatValue/(PayableAmount - VatValue) *100,0) AS TaxCategoryPercent"
   
   s = s & " ,ROUND(VatValue / NULLIF((PayableAmount - VatValue), 0) * 100, 0) AS TaxCategoryPercent"

   s = s & " ,tblEInvoice.last_changed"
   
   s = s & " ,tblEInvoice.PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
   s = s & " ,10 AS DocType ,VatValue,"
   
   
    
's = s & " ISNULL("
's = s & "         CASE"
's = s & "             WHEN ISNULL(ROUND(VatValue / NULLIF((PayableAmount - VatValue), 0) * 100, 0), 0) = 0 THEN 1"
's = s & "             WHEN tblEInvoice.Export = 1 THEN 1"
's = s & "             WHEN (tblEInvoice.Export Is Null Or tblEInvoice.Export = 0)"
's = s & "                  AND ISNULL(ROUND(VatValue / NULLIF((PayableAmount - VatValue), 0) * 100, 0), 0) = 0 THEN 1"
's = s & "             Else tblEInvoice.Export"
's = s & "         END,"
's = s & "     0) AS Export"
    
    
s = s & "             tblEInvoice.Export"
    s = s & " From "
s = s & " tblEInvoice"

s = s & " left outer JOIN dbo.transactionsVatDetails"
s = s & " ON tblEInvoice.Transaction_ID = dbo.transactionsVatDetails.Transaction_ID and transactionsVatDetails.TableName = 'tblEInvoice'"
s = s & "         AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"
s = s & " Where 1 = 1"


    
  
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate <=" & SQLDate(ToDate, True) & " "
    End If
    
    s = s & " and ISNULL(tblEInvoice.zatcaStatus, 0) <> 1"
       
      If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     tblEInvoice.branch_id In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
  
  
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       tblEInvoice.branch_id = " & val(DcBranches(3).BoundText)
        
        
    End If
    
    
     s = s & "  Union all"
    
    s = s & " SELECT tblActivitesType.id as ActivityTypeId,'Service Invoice' as TypeName, TblBranchesData.branch_id,TblBranchesData.branch_name BranchName, tblActivitesType.Name as ActivityName,"
    s = s & "     notes_all.Invoicetype"
   s = s & " ,dbo.notes_all.ErrorMessageS,chkTaxExempt = 0"
   s = s & " ,dbo.notes_all.NoteDate DateBaptizing"
   s = s & " ,CAST(dbo.notes_all.order_no AS VARCHAR(10)) order_no"
   s = s & " ,order_no ReturnSerial"
   s = s & " ,dbo.notes_all.NoteDate SalesInvoiceDate"
   s = s & " ,dbo.notes_all.Noteid Transaction_ID"
   s = s & " ,PayeeFinancialAccount = (SELECT"
    s = s & "             IBan"
    s = s & "      From BanksData"
    s = s & "         WHERE bankid = 0)"
   s = s & " ,dbo.notes_all.Noteid AS Expr1"
   s = s & " ,Transaction_Type =  notes_all.NoteType "
   
    s = s & " ,CAST(CAST(dbo.notes_all.NoteSerial1 AS BIGINT) AS NVARCHAR(50)) AS id"

   
   s = s & " ,dbo.notes_all.NoteDate AS IssueDate"
   s = s & " ,dbo.notes_all.RecTime AS IssueTim"
   s = s & " ,dbo.notes_all.InvoiceTypeCodeID"
   s = s & " ,dbo.notes_all.InvoiceTypeCodename"
   s = s & " ,dbo.notes_all.DocumentCurrencyCode"
   s = s & " ,dbo.notes_all.TaxCurrencyCode"
   s = s & " ,dbo.notes_all.InvoiceDocumentReferenceID"
   s = s & " ,dbo.notes_all.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.notes_all.ActualDeliveryDate"
   s = s & " ,dbo.notes_all.LatestDeliveryDate"
   s = s & " ,dbo.notes_all.PaymentMeansCode"
   s = s & " ,dbo.notes_all.InstructionNote"
   s = s & " ,dbo.notes_all.paymentnote"
   s = s & " ,dbo.TblCustemers.CustGID AS Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.TblCustemers.StreetName"
   s = s & " ,dbo.TblCustemers.AdditionalStreetName"
   s = s & " ,dbo.TblCustemers.BuildingNumber"
   s = s & " ,dbo.TblCustemers.PlotIdentification"
   s = s & " ,dbo.TblCustemers.CityName"
   s = s & " ,dbo.TblCustemers.PostalZone"
   s = s & " ,dbo.TblCustemers.CountrySubentity"
   s = s & " ,dbo.TblCustemers.CitySubdivisionName"
   s = s & " ,dbo.TblCustemers.IdentificationCode"
   s = s & " ,dbo.TblCustemers.CusNamee AS RegistrationName"
   s = s & " ,dbo.TblCustemers.VATNO AS CompanyID, dbo.TblCustemers.Id700 "
   s = s & " ,0 AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,'S' AS TaxCategoryID"
   s = s & " ,'15' AS TaxCategoryPercent"
   s = s & " ,notes_all.last_changed"
   
   s = s & " , Note_Value AS PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
   s = s & " ,4 AS DocType,0 VatValue,TblCustemers.Export"
    s = s & " From dbo.TblCustemers"
s = s & " Right outer JOIN dbo.notes_all"
s = s & "     ON dbo.TblCustemers.CusID = dbo.notes_all.CusID"
s = s & " left outer JOIN dbo.transactionsVatDetails"
s = s & " ON dbo.notes_all.NoteID = dbo.transactionsVatDetails.Transaction_ID and transactionsVatDetails.TableName = 'notes_all'"
s = s & "         AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"

s = s & " Inner join TblBranchesData On TblBranchesData.branch_id = notes_all.branch_no inner join tblActivitesType On tblActivitesType.id =TblBranchesData.ActivityTypeId  "

s = s & " Where 1 = 1"
s = s & " AND notes_all.NoteType IN (85)"

    
  
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.notes_all.NoteDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.notes_all.NoteDate <=" & SQLDate(ToDate, True) & " "
    End If
    
    s = s & " and ISNULL(notes_all.zatcaStatus, 0) <> 1"
   ' s = s & "                      ORDER by  InvoiceTypeCodeID desc ,Transaction_Date, notes_allerial1"
    
    
          
    If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     notes_all.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
       
    
            
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       notes_all.branch_no= " & val(DcBranches(3).BoundText)
        
        
    End If
    
    
      s = s & "  Union all"
    s = s & " SELECT  tblActivitesType.id as ActivityTypeId,'Projects Invoice' as TypeName, TblBranchesData.branch_id,TblBranchesData.branch_name BranchName, tblActivitesType.Name as ActivityName,"
    s = s & " project_billl.Invoicetype"
   s = s & " ,dbo.project_billl.ErrorMessageS,chkTaxExempt = 0"
   s = s & " ,dbo.project_billl.bill_date DateBaptizing"
   s = s & " ,cast (dbo.project_billl.order_no    as VARCHAR(10)) order_no"
   s = s & " ,'' ReturnSerial"
   s = s & " ,dbo.project_billl.bill_date SalesInvoiceDate"
   s = s & " ,dbo.project_billl.id Transaction_ID"
   s = s & " ,PayeeFinancialAccount = (SELECT"
    s = s & "             IBan"
    s = s & "         From BanksData"
    s = s & "         WHERE bankid = 0)"
   s = s & " ,dbo.project_billl.ID AS Expr1"
   s = s & " ,1 as Transaction_Type"
   
   s = s & " ,CAST(dbo.project_billl.NoteSerial1 AS NVARCHAR(50)) AS id"
   s = s & " ,dbo.project_billl.bill_date AS IssueDate"
   s = s & " ,dbo.project_billl.RecTime AS IssueTim"
   s = s & " ,dbo.project_billl.InvoiceTypeCodeID"
   s = s & " ,dbo.project_billl.InvoiceTypeCodename"
   s = s & " ,dbo.project_billl.DocumentCurrencyCode"
   s = s & " ,dbo.project_billl.TaxCurrencyCode"
   s = s & " ,dbo.project_billl.InvoiceDocumentReferenceID"
   s = s & " ,dbo.project_billl.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.project_billl.ActualDeliveryDate"
   s = s & " ,dbo.project_billl.LatestDeliveryDate"
   s = s & " ,dbo.project_billl.PaymentMeansCode"
   s = s & " ,dbo.project_billl.InstructionNote"
   s = s & " ,dbo.project_billl.paymentnote"
   s = s & " ,dbo.TblCustemers.CustGID AS Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.TblCustemers.StreetName"
   s = s & " ,dbo.TblCustemers.AdditionalStreetName"
   s = s & " ,dbo.TblCustemers.BuildingNumber"
   s = s & " ,dbo.TblCustemers.PlotIdentification"
   s = s & " ,dbo.TblCustemers.CityName"
   s = s & " ,dbo.TblCustemers.PostalZone"
   s = s & " ,dbo.TblCustemers.CountrySubentity"
   s = s & " ,dbo.TblCustemers.CitySubdivisionName"
   s = s & " ,dbo.TblCustemers.IdentificationCode"
   s = s & " ,dbo.TblCustemers.CusNamee AS RegistrationName"
   s = s & " ,dbo.TblCustemers.VATNO AS CompanyID, dbo.TblCustemers.Id700 "
   
   's = s & " ,project_billl.DiscountGMater + project_billl.Discount4 + project_billl.advancedPayment AS allowancechargeAmount"
    's = s & " ,project_billl.discount + project_billl.DiscountGMater +  project_billl.advancedPayment AS allowancechargeAmount"
    s = s & " ,project_billl.Discount4  AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,'S' AS TaxCategoryID"
   s = s & " ,'15' AS TaxCategoryPercent"
   s = s & " ,project_billl.last_changed"
   
   s = s & " ,project_billl.total+ project_billl.FATValue AS PayableAmount"
  ' s = s & " ,project_billl.total AS PayableAmount"
   's = s & " ,project_billl.TotalValue  AS PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
   s = s & " ,1 AS DocType ,0 VatValue,TblCustemers.Export"


    s = s & " From dbo.TblCustemers"
    s = s & " INNER JOIN dbo.project_billl"
    s = s & " inner join projects On project_no = projects.id"
    s = s & " ON dbo.TblCustemers.CusID = dbo.projects.End_user_id"
    s = s & " LEFT OUTER JOIN dbo.transactionsVatDetails"
    s = s & " ON dbo.project_billl.ID = dbo.transactionsVatDetails.Transaction_ID"
    s = s & " and transactionsVatDetails.TableName = 'project_billl'"
    s = s & "         AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"

    s = s & " Inner join TblBranchesData On TblBranchesData.branch_id = project_billl.Branch_NO inner join tblActivitesType On tblActivitesType.id =TblBranchesData.ActivityTypeId  "
    
    
    s = s & " Where 1 = 1 and project_billl.bill_type = 0"
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.project_billl.bill_date >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.project_billl.bill_date <=" & SQLDate(ToDate, True) & " "
    End If
    
   If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     project_billl.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
        
            
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and      project_billl.branch_no = " & val(DcBranches(3).BoundText)
        
        
    End If
            
    
    s = s & " and ISNULL(project_billl.zatcaStatus, 0) <> 1"
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
       
     s = s & "  Union all"
    
    s = s & " SELECT tblActivitesType.id as ActivityTypeId,'«ÃÊ— Ìœ' as TypeName,  TblBranchesData.branch_id,TblBranchesData.branch_name BranchName, tblActivitesType.Name as ActivityName, "
    s = s & "     TblHandWages.Invoicetype"
   s = s & " ,dbo.TblHandWages.ErrorMessageS,chkTaxExempt = 0"
   s = s & " ,dbo.TblHandWages.RecordDate DateBaptizing"
   s = s & " ,CAST(dbo.TblHandWages.NoteSerial1 AS VARCHAR(10)) order_no"
   s = s & " ,order_no ReturnSerial"
   s = s & " ,dbo.TblHandWages.RecordDate SalesInvoiceDate"
   s = s & " ,dbo.TblHandWages.id Transaction_ID"
   s = s & " ,PayeeFinancialAccount = (SELECT"
    s = s & "             IBan"
    s = s & "      From BanksData"
    s = s & "         WHERE bankid = 0)"
   s = s & " ,dbo.TblHandWages.Noteid AS Expr1"
    s = s & " ,10 as Transaction_Type"

   
   s = s & " ,CAST(dbo.TblHandWages.NoteSerial1 AS NVARCHAR(50)) AS id"
   s = s & " ,dbo.TblHandWages.RecordDate AS IssueDate"
   s = s & " ,dbo.TblHandWages.RecTime AS IssueTim"
   s = s & " ,dbo.TblHandWages.InvoiceTypeCodeID"
   s = s & " ,dbo.TblHandWages.InvoiceTypeCodename"
   s = s & " ,dbo.TblHandWages.DocumentCurrencyCode"
   s = s & " ,dbo.TblHandWages.TaxCurrencyCode"
   s = s & " ,dbo.TblHandWages.InvoiceDocumentReferenceID"
   s = s & " ,dbo.TblHandWages.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.TblHandWages.ActualDeliveryDate"
   s = s & " ,dbo.TblHandWages.LatestDeliveryDate"
   s = s & " ,dbo.TblHandWages.PaymentMeansCode"
   s = s & " ,dbo.TblHandWages.InstructionNote"
   s = s & " ,dbo.TblHandWages.paymentnote"
   s = s & " ,dbo.TblCustemers.CustGID AS Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.TblCustemers.StreetName"
   s = s & " ,dbo.TblCustemers.AdditionalStreetName"
   s = s & " ,dbo.TblCustemers.BuildingNumber"
   s = s & " ,dbo.TblCustemers.PlotIdentification"
   s = s & " ,dbo.TblCustemers.CityName"
   s = s & " ,dbo.TblCustemers.PostalZone"
   s = s & " ,dbo.TblCustemers.CountrySubentity"
   s = s & " ,dbo.TblCustemers.CitySubdivisionName"
   s = s & " ,dbo.TblCustemers.IdentificationCode"
   s = s & " ,dbo.TblCustemers.CusName AS RegistrationName"
   s = s & " ,dbo.TblCustemers.VATNO AS CompanyID , dbo.TblCustemers.Id700 "
   s = s & " ,0 AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,'S' AS TaxCategoryID"
   s = s & " ,'15' AS TaxCategoryPercent"
   s = s & " ,TblHandWages.last_changed"
   
   s = s & " ,TotalNet AS PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
   s = s & " ,5 AS DocType,vat2 as VatValue,TblCustemers.Export"
'    s = s & " From dbo.TblCustemers"
's = s & " INNER JOIN dbo.TblHandWages"
's = s & "     ON dbo.TblCustemers.CusID = dbo.TblHandWages.CusID"
's = s & " left outer JOIN dbo.transactionsVatDetails"
's = s & " ON dbo.TblHandWages.ID = dbo.transactionsVatDetails.Transaction_ID and transactionsVatDetails.TableName = 'TblHandWages'"
's = s & "         AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"
's = s & " Inner join TblBranchesData On TblBranchesData.branch_id = TblHandWages.BranchId inner join tblActivitesType On tblActivitesType.id =TblBranchesData.ActivityTypeId  "

s = s & " From transactionsVatDetails"
s = s & " RIGHT OUTER JOIN TblHandWages"
s = s & "     ON transactionsVatDetails.Transaction_ID = TblHandWages.ID"
s = s & "     and transactionsVatDetails.TableName = 'TblHandWages'"
s = s & "         AND ISNULL(transactionsVatDetails.IsDeleted, 0) = 0"
s = s & " LEFT OUTER JOIN TblCustemers"

s = s & "         on TblHandWages.CusID = TblCustemers.CusID"
s = s & " LEFT OUTER JOIN TblBranchesData"
s = s & "     ON TblBranchesData.branch_id = TblHandWages.BranchID"
s = s & " LEFT OUTER JOIN tblActivitesType"
s = s & "     ON tblActivitesType.id = TblBranchesData.ActivityTypeId"

s = s & " Where 1 = 1"


    
  
    
    If Not IsNull(FromDate.value) Then
       s = s & " and dbo.TblHandWages.recordDate >=" & SQLDate(FromDate, True) & " "
    End If
    
    If Not IsNull(ToDate.value) Then
       s = s & " and dbo.TblHandWages.recordDate <=" & SQLDate(ToDate, True) & " "
    End If
    
    s = s & " and ISNULL(TblHandWages.zatcaStatus, 0) <> 1"
   ' s = s & "                      ORDER by  InvoiceTypeCodeID desc ,Transaction_Date, NoteSerial1"
    
    
        
    If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     TblHandWages.BranchId In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If

      
  
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       TblHandWages.BranchId= " & val(DcBranches(3).BoundText)
    End If
    
    
    s = s & " and IsNull(TblHandWages.IsHiddenVat,0) = 0        "
    
    

s = s & "  Union all"

s = s & " SELECT tblActivitesType.id as ActivityTypeId, 'RS Contract' as TypeName, TblBranchesData.branch_id, TblBranchesData.branch_name AS BranchName, tblActivitesType.Name AS ActivityName, "
s = s & "     tblContractInsAllocationsDetails.Invoicetype, "
s = s & "     tblContractInsAllocationsDetails.ErrorMessageS, chkTaxExempt = 0, "
s = s & "     tblContractInsAllocationsDetails.DateRec AS DateBaptizing, "
s = s & "     CAST(tblContractInsAllocationsDetails.NoteSerial1H AS VARCHAR(10)) AS order_no, "
s = s & "     '' AS ReturnSerial, "
s = s & "     tblContractInsAllocationsDetails.Installdate AS SalesInvoiceDate, "
s = s & "     tblContractInsAllocationsDetails.id AS Transaction_ID, "
s = s & "     PayeeFinancialAccount = (SELECT IBan FROM BanksData WHERE BankID = 0), "
s = s & "     tblContractInsAllocationsDetails.Id AS Expr1, 10 AS Transaction_Type, "
s = s & "     tblContractInsAllocationsDetails.NoteSerial1H AS id, "
s = s & "     tblContractInsAllocationsDetails.DateRec AS IssueDate, "
s = s & "     tblContractInsAllocationsDetails.RecTime AS IssueTim, "
s = s & "     tblContractInsAllocationsDetails.InvoiceTypeCodeID, "
s = s & "     tblContractInsAllocationsDetails.InvoiceTypeCodename, "
s = s & "     tblContractInsAllocationsDetails.DocumentCurrencyCode, "
s = s & "     tblContractInsAllocationsDetails.TaxCurrencyCode, "
s = s & "     tblContractInsAllocationsDetails.InvoiceDocumentReferenceID, "
s = s & "     tblContractInsAllocationsDetails.AdditionalDocumentReferenceICVUUID, "
s = s & "     tblContractInsAllocationsDetails.ActualDeliveryDate, "
s = s & "     tblContractInsAllocationsDetails.LatestDeliveryDate, "
s = s & "     tblContractInsAllocationsDetails.PaymentMeansCode, "
s = s & "     tblContractInsAllocationsDetails.InstructionNote, "
s = s & "     tblContractInsAllocationsDetails.paymentnote, "
s = s & "     TblCustemers.CustGID AS Identificationid, "
s = s & "     'CRN' AS schemeID, "
s = s & "     TblCustemers.StreetName, "
s = s & "     TblCustemers.AdditionalStreetName, "
s = s & "     TblCustemers.BuildingNumber, "
s = s & "     TblCustemers.PlotIdentification, "
s = s & "     TblCustemers.CityName, "
s = s & "     TblCustemers.PostalZone, "
s = s & "     TblCustemers.CountrySubentity, "
s = s & "     TblCustemers.CitySubdivisionName, "
s = s & "     TblCustemers.IdentificationCode, "
s = s & "     TblCustemers.CusName AS RegistrationName, "
s = s & "     TblCustemers.VATNO AS CompanyID, TblCustemers.Id700, "
s = s & "     0 AS allowancechargeAmount, "
s = s & "     'Discount' AS AllowanceChargeReason, "

s = s & "     CASE WHEN FATYou = 0 THEN 'O' ELSE 'S' END AS TaxCategoryID, "
s = s & "     FATYou AS TaxCategoryPercent, "


s = s & "     tblContractInsAllocationsDetails.last_changed, "
s = s & "     tblContractInsAllocationsDetails.installValue + isnull(VATValue,0) AS PayableAmount, "
s = s & "     0 AS PrepaidAmount, "
s = s & "     transactionsVatDetails.SingedXMLFileName, "
s = s & "     transactionsVatDetails.PIH, "
s = s & "     transactionsVatDetails.QRCode, "
s = s & "     transactionsVatDetails.UUID, "
s = s & "     transactionsVatDetails.InvoiceHash, "
s = s & "     transactionsVatDetails.EncodedInvoice, "
s = s & "     transactionsVatDetails.SingedXML, "
s = s & "     transactionsVatDetails.QrCodeDataPath, "
s = s & "     6 AS DocType, FATValue AS VatValue, TblCustemers.Export "

s = s & "    From tblActivitesType"
s = s & "    RIGHT OUTER JOIN TblContract"
s = s & "    LEFT OUTER JOIN TblBranchesData"
s = s & "        ON TblContract.Branch_NO = TblBranchesData.branch_id"
s = s & "        ON tblActivitesType.id = TblBranchesData.ActivityTypeId"
s = s & "    RIGHT OUTER JOIN tblContractInsAllocationsDetails"
s = s & "    INNER JOIN TblCustemers"
s = s & "        ON tblContractInsAllocationsDetails.CusID = TblCustemers.CusID"
s = s & "        ON TblContract.ContNo = tblContractInsAllocationsDetails.ContNo"
s = s & "    LEFT OUTER JOIN transactionsVatDetails"
s = s & "        ON tblContractInsAllocationsDetails.id = transactionsVatDetails.Transaction_ID"
s = s & "            AND transactionsVatDetails.TableName = 'tblContractInsAllocationsDetails'"
s = s & "            AND ISNULL(transactionsVatDetails.IsDeleted,"
s = s & "            0) = 0"

s = s & " WHERE 1 = 1 "

If Not IsNull(FromDate.value) Then
    s = s & " AND tblContractInsAllocationsDetails.Installdate >= " & SQLDate(FromDate, True) & " "
End If

If Not IsNull(ToDate.value) Then
    s = s & " AND tblContractInsAllocationsDetails.Installdate <= " & SQLDate(ToDate, True) & " "
End If

s = s & " AND ISNULL(tblContractInsAllocationsDetails.zatcaStatus, 0) <> 1 "

If Trim(DcBranches(2).Text) <> "" Then
    s = s & " AND TblBranchesData.branch_id IN (SELECT branch_id FROM TblBranchesData WHERE ActivityTypeId = " & DcBranches(2).BoundText & ") "
End If

If Trim(DcBranches(3).Text) <> "" Then
    s = s & " AND TblBranchesData.branch_id = " & val(DcBranches(3).BoundText) & " "
End If








    
    s = s & "                      ORDER by  InvoiceTypeCodeID desc ,Transaction_Date, NoteSerial1"
    Set rsDummy = New ADODB.Recordset
    
    
    If SystemOptions.CanUploadZakatOpt And chkNotes.value = 0 Then
        
        Dim rs As New ADODB.Recordset
        
        Dim SerialKey As String
        Dim BranchID As Integer
        Dim dueDate As Date
        ConectionFirst

s = ""
s = s & " ;WITH LastSerial AS ("
s = s & "   SELECT"
s = s & "     pc.DepartmentId AS branch_id,"
s = s & "     RIGHT(CAST(YEAR(pcb.BatchDate) AS varchar(4)), 2) AS y,"
s = s & "     RIGHT('00' + CAST(MONTH(pcb.BatchDate) AS varchar), 2) AS m,"
s = s & "     MAX(CAST(RIGHT(ISNULL(p.NoteSerial1H, '0'), 4) AS int)) AS LastNo"
s = s & "   FROM PropertyDueBatchDetail p"
s = s & "   INNER JOIN PropertyContractBatch pcb ON pcb.Id = p.PropertyContractBatchId"
s = s & "   INNER JOIN PropertyContract pc ON pc.Id = pcb.MainDocId"
s = s & "   LEFT OUTER JOIN Department d ON pc.DepartmentId = d.Id"
s = s & "   WHERE ISNULL(p.NoteSerial1H,'') <> ''"
s = s & "   GROUP BY pc.DepartmentId, RIGHT(CAST(YEAR(pcb.BatchDate) AS varchar(4)), 2), RIGHT('00' + CAST(MONTH(pcb.BatchDate) AS varchar), 2)"
s = s & " ),"
s = s & " SerialBase AS ("
s = s & "   SELECT"
s = s & "     p.ID,"
s = s & "     pc.DepartmentId AS branch_id,"
s = s & "     RIGHT(CAST(YEAR(pcb.BatchDate) AS varchar(4)), 2) AS y,"
s = s & "     RIGHT('00' + CAST(MONTH(pcb.BatchDate) AS varchar), 2) AS m,"
s = s & "     CAST(pc.DepartmentId AS varchar(10)) +"
s = s & "     RIGHT(CAST(YEAR(pcb.BatchDate) AS varchar(4)), 2) +"
s = s & "     RIGHT('00' + CAST(MONTH(pcb.BatchDate) AS varchar), 2) AS SerialKey,"
s = s & "     ROW_NUMBER() OVER ("
s = s & "         PARTITION BY pc.DepartmentId, YEAR(pcb.BatchDate), MONTH(pcb.BatchDate)"
s = s & "         ORDER BY pcb.BatchDate, p.ID"
s = s & "     ) AS RowNo"
s = s & "   FROM PropertyDueBatchDetail p"
s = s & "   INNER JOIN PropertyContractBatch pcb ON pcb.Id = p.PropertyContractBatchId"
s = s & "   INNER JOIN PropertyContract pc ON pc.Id = pcb.MainDocId"
s = s & "   LEFT OUTER JOIN Department d ON pc.DepartmentId = d.Id"
s = s & "   WHERE ISNULL(p.IsSelected,0)=1"
s = s & "     AND ISNULL(p.zatcaStatus,0)<>1"
s = s & "     AND ISNULL(p.NoteSerial1H,'') = ''"
's = s & "     and pc.id not in (Select PropertyContractTermination.PropertyContractId from PropertyContractTermination)"
If Not IsNull(FromDate.value) Then
    s = s & " AND pcb.BatchDate >= '" & Format(FromDate.value, "yyyy-MM-dd") & "'"
End If
If Not IsNull(ToDate.value) Then
    s = s & " AND pcb.BatchDate <= '" & Format(ToDate.value, "yyyy-MM-dd") & "'"
End If
s = s & " )"
s = s & " UPDATE p"
s = s & " SET NoteSerial1H = s.SerialKey + RIGHT('0000' + CAST(ISNULL(l.LastNo, 0) + s.RowNo AS varchar), 4)"
s = s & " FROM PropertyDueBatchDetail p"
s = s & " JOIN SerialBase s ON p.ID = s.ID"
s = s & " LEFT JOIN LastSerial l ON s.branch_id = l.branch_id AND s.y = l.y AND s.m = l.m"


' ??? ?????????
POSConnection.Execute s

        
        '=== Helper to append one SQL line safely ===


'=== Build SQL ===
 s = ""

'===== CTE: Base_RS =====
SqlAdd s, "WITH Base_RS AS ("
SqlAdd s, "  SELECT"
SqlAdd s, "    CAST(0 AS int)                         AS ActivityTypeId,"
SqlAdd s, "    N'RS Contract'                         AS TypeName,"
SqlAdd s, "    dep.Id                                  AS branch_id,"
SqlAdd s, "    dep.ArName                              AS BranchName,"
SqlAdd s, "    N''                                     AS ActivityName,"
SqlAdd s, "    ISNULL(CAST(388 AS int),388)            AS Invoicetype,"
SqlAdd s, "    ISNULL(pdbd.ErrorMessageS,N'')          AS ErrorMessageS,"
SqlAdd s, "    CAST(0 AS bit)                           AS chkTaxExempt,"
SqlAdd s, "    pcb.BatchDate                            AS DateBaptizing,"
SqlAdd s, "    N''                                      AS order_no,"
SqlAdd s, "    N''                                      AS ReturnSerial,"
SqlAdd s, "    pcb.BatchDate                            AS SalesInvoiceDate,"
SqlAdd s, "    CAST(pdbd.Id AS bigint)                  AS Transaction_ID,"
SqlAdd s, "    N''                                      AS PayeeFinancialAccount,"
SqlAdd s, "    CAST(pdbd.Id AS bigint)                  AS Expr1,"
SqlAdd s, "    CAST(10 AS int)                          AS Transaction_Type,"
SqlAdd s, "    CAST(pdbd.NoteSerial1H AS nvarchar(50))  AS id,"
SqlAdd s, "    pcb.BatchDate                            AS IssueDate,"
SqlAdd s, "    ISNULL(CONVERT(nvarchar(8),GETDATE(),108),N'') AS IssueTim,"
SqlAdd s, "    ISNULL(pdbd.InvoiceTypeCodeID,388)       AS InvoiceTypeCodeID,"

SqlAdd s, "         InvoiceTypeCodename,"
SqlAdd s, "    ISNULL(pdbd.DocumentCurrencyCode,N'SAR') AS DocumentCurrencyCode,"
SqlAdd s, "    ISNULL(pdbd.TaxCurrencyCode,N'SAR')      AS TaxCurrencyCode,"
SqlAdd s, "    ISNULL(pdbd.InvoiceDocumentReferenceID,N'')     AS InvoiceDocumentReferenceID,"
SqlAdd s, "    ISNULL(pdbd.AdditionalDocumentReferenceICVUUID,N'') AS AdditionalDocumentReferenceICVUUID,"
SqlAdd s, "    ISNULL(pdbd.ActualDeliveryDate,CAST(GETDATE() AS date))  AS ActualDeliveryDate,"
SqlAdd s, "    ISNULL(pdbd.LatestDeliveryDate,CAST(GETDATE() AS date))  AS LatestDeliveryDate,"
SqlAdd s, "    ISNULL(CAST(pdbd.PaymentMeansCode AS float),30)          AS PaymentMeansCode,"
SqlAdd s, "    ISNULL(pdbd.InstructionNote,N'')                 AS InstructionNote,"
SqlAdd s, "    ISNULL(pdbd.paymentnote,N'Payment by Credit')    AS paymentnote,"
SqlAdd s, "    pr.RegistrationNo                       AS Identificationid,"
SqlAdd s, "    N'CRN'                                   AS schemeID,"
SqlAdd s, "    pr.Address                               AS StreetName,"
SqlAdd s, "    pr.Address                               AS AdditionalStreetName,"
SqlAdd s, "    NULL                                     AS BuildingNumber,"
SqlAdd s, "    NULL                                     AS PlotIdentification,"
SqlAdd s, "    N'Riyadh'                                AS CityName,"
SqlAdd s, "    N'12345'                                 AS PostalZone,"
SqlAdd s, "    NULL                                     AS CountrySubentity,"
SqlAdd s, "    NULL                                     AS CitySubdivisionName,"
SqlAdd s, "    NULL                                     AS IdentificationCode,"
SqlAdd s, "    pr.ArName                                AS RegistrationName,"
SqlAdd s, "    pr.VATNo                                 AS CompanyID,"
SqlAdd s, "    NULL                                     AS Id700,"
SqlAdd s, "    CAST(0 AS float)                         AS allowancechargeAmount,"
SqlAdd s, "    N'Discount'                              AS AllowanceChargeReason,"
SqlAdd s, "    CASE WHEN ISNULL(pc.VATPercentage,0)=0 THEN 'O' ELSE 'S' END AS TaxCategoryID,"
SqlAdd s, "    ISNULL(pc.VATPercentage,0)*100           AS TaxCategoryPercent,"
SqlAdd s, "    pdbd.last_changed                        AS last_changed,"
SqlAdd s, "    pcb.BatchTotal                            AS PayableAmount,"
SqlAdd s, "    CAST(0 AS float)                          AS PrepaidAmount,"
SqlAdd s, "    ISNULL(tvd.SingedXMLFileName,N'')        AS SingedXMLFileName,"
SqlAdd s, "    ISNULL(tvd.PIH,N'')                      AS PIH,"
SqlAdd s, "    ISNULL(tvd.QRCode,N'')                   AS QRCode,"
SqlAdd s, "    ISNULL(tvd.UUID,N'')                     AS UUID,"
SqlAdd s, "    ISNULL(tvd.InvoiceHash,N'')              AS InvoiceHash,"
SqlAdd s, "    ISNULL(tvd.EncodedInvoice,N'')           AS EncodedInvoice,"
SqlAdd s, "    ISNULL(tvd.SingedXML,N'')                AS SingedXML,"
SqlAdd s, "    ISNULL(tvd.QrCodeDataPath,N'')           AS QrCodeDataPath,"
SqlAdd s, "    CAST(6 AS int)                           AS DocType,"
SqlAdd s, "    CAST(ROUND("
SqlAdd s, "        ISNULL(pcb.BatchRentValueTaxes,0)+ISNULL(pcb.BatchWaterValueTaxes,0)+"
SqlAdd s, "        ISNULL(pcb.BatchElectricityValueTaxes,0)+ISNULL(pcb.BatchCommissionValueTaxes,0)+"
SqlAdd s, "        ISNULL(pcb.BatchGasValueTaxes,0)+ISNULL(pcb.BatchServicesValueTaxes,0)+"
SqlAdd s, "        ISNULL(pcb.BatchInsuranceValueTaxes,0),2) AS decimal(18,2)) AS VatValue,"
SqlAdd s, "    ISNULL(pdbd.Export,0)                    AS Export,"
SqlAdd s, "    STUFF(("
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchRentValueTaxes,0)>0 THEN N' | ÷—Ì»… «·≈ÌÃ«— ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchRentValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchWaterValueTaxes,0)>0 THEN N' | ÷—Ì»… «·„Ì«Â ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchWaterValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchElectricityValueTaxes,0)>0 THEN N' | ÷—Ì»… «·ﬂÂ—»«¡ ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchElectricityValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchCommissionValueTaxes,0)>0 THEN N' | ÷—Ì»… «·”⁄Ì ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchCommissionValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchGasValueTaxes,0)>0 THEN N' | ÷—Ì»… «·€«“ ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchGasValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchServicesValueTaxes,0)>0 THEN N' | ÷—Ì»… «·Œœ„«  ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchServicesValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchInsuranceValueTaxes,0)>0 THEN N' | ÷—Ì»… «· √„Ì‰ ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchInsuranceValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END)"
SqlAdd s, "    ),1,3,N'') AS VatBreakdownNote"
SqlAdd s, "  FROM PropertyDueBatchDetail pdbd"
SqlAdd s, "  JOIN PropertyContractBatch pcb ON pcb.Id = pdbd.PropertyContractBatchId"
SqlAdd s, "  JOIN PropertyContract pc       ON pc.Id  = pcb.MainDocId"
SqlAdd s, "  JOIN PropertyRenter pr         ON pc.PropertyRenterId = pr.Id"
SqlAdd s, "  LEFT JOIN Department dep       ON pc.DepartmentId = dep.Id"
SqlAdd s, "  LEFT JOIN transactionsVatDetails tvd"
SqlAdd s, "         ON tvd.Transaction_ID = pdbd.Id"
SqlAdd s, "        AND tvd.TableName      = 'PropertyDueBatchDetail'"
SqlAdd s, "        AND ISNULL(tvd.IsDeleted,0)=0"
SqlAdd s, "  WHERE ISNULL(pdbd.IsSelected,0)=1"
SqlAdd s, "    AND ISNULL(pdbd.zatcaStatus,0)<>1"
SqlAdd s, " and pc.id Not in (Select PropertyContractId from PropertyContractTermination)"
'ó ›·« — «· «—ÌŒ ·‹ RS
If Not IsNull(FromDate.value) Then SqlAdd s, "    AND pcb.BatchDate >= " & SQLDate(FromDate, True)
If Not IsNull(ToDate.value) Then SqlAdd s, "    AND pcb.BatchDate <= " & SQLDate(ToDate, True)

SqlAdd s, ")"

'===== CTE: Base_NOTE =====
SqlAdd s, ", Base_NOTE AS ("
SqlAdd s, "  SELECT"
SqlAdd s, "    ISNULL(dep.ActivityId,0)                  AS ActivityTypeId,"
SqlAdd s, "    N'Credit Or Debit Note (WEB)'             AS TypeName,"
SqlAdd s, "    dep.Id                                    AS branch_id,"
SqlAdd s, "    dep.ArName                                AS BranchName,"
SqlAdd s, "    N''                                       AS ActivityName,"
SqlAdd s, "     Invoicetype,"
SqlAdd s, "    ISNULL(d.ErrorMessageS,N'')               AS ErrorMessageS,"
SqlAdd s, "    CAST(0 AS bit)                            AS chkTaxExempt,"
SqlAdd s, "    d.[Date]                                  AS DateBaptizing,"
SqlAdd s, "    CAST(ISNULL(d.DocumentNumber,N'') AS nvarchar(50)) AS order_no,"
SqlAdd s, "    N''                                       AS ReturnSerial,"
SqlAdd s, "    d.[Date]                                  AS SalesInvoiceDate,"
SqlAdd s, "    CAST(d.Id AS bigint)                      AS Transaction_ID,"
SqlAdd s, "    N''                                       AS PayeeFinancialAccount,"
SqlAdd s, "    CAST(d.Id AS bigint)                      AS Expr1,"
SqlAdd s, "    CASE d.DebitAndCreditNotificationTypeId WHEN 4 THEN 383 WHEN 5 THEN 381 ELSE 0 END AS Transaction_Type,"
SqlAdd s, "    ISNULL(d.InvoiceNo,ISNULL(d.DocumentNumber,N'')) AS id,"
SqlAdd s, "    d.[Date]                                  AS IssueDate,"
SqlAdd s, "    ISNULL(d.RecTime,CONVERT(nvarchar(8),GETDATE(),108)) AS IssueTim,"
SqlAdd s, "    ISNULL(CAST(d.InvoiceTypeCodeID AS int),0) AS InvoiceTypeCodeID,"
SqlAdd s, "     InvoiceTypeCodename,"
SqlAdd s, "    ISNULL(d.DocumentCurrencyCode,N'SAR')     AS DocumentCurrencyCode,"
SqlAdd s, "    ISNULL(d.TaxCurrencyCode,N'SAR')          AS TaxCurrencyCode,"
SqlAdd s, "    ISNULL(d.InvoiceDocumentReferenceID,N'')  AS InvoiceDocumentReferenceID,"
SqlAdd s, "    ISNULL(d.AdditionalDocumentReferenceICVUUID,N'') AS AdditionalDocumentReferenceICVUUID,"
SqlAdd s, "    ISNULL(d.ActualDeliveryDate,CAST(GETDATE() AS date)) AS ActualDeliveryDate,"
SqlAdd s, "    ISNULL(d.LatestDeliveryDate,CAST(GETDATE() AS date)) AS LatestDeliveryDate,"
SqlAdd s, "    ISNULL(CAST(d.PaymentMeansCode AS float),30)        AS PaymentMeansCode,"
SqlAdd s, "    ISNULL(d.InstructionNote,N'')            AS InstructionNote,"
SqlAdd s, "    ISNULL(d.paymentnote,N'Payment by Credit') AS paymentnote,"
SqlAdd s, "    COALESCE(r.RegistrationNo,c.RegistrationNo,v.RegistrationNo,N'') AS Identificationid,"
SqlAdd s, "    N'CRN'                                     AS schemeID,"
SqlAdd s, "    N''                                        AS StreetName,"
SqlAdd s, "    N''                                        AS AdditionalStreetName,"
SqlAdd s, "    NULL                                       AS BuildingNumber,"
SqlAdd s, "    NULL                                       AS PlotIdentification,"
SqlAdd s, "    N''                                        AS CityName,"
SqlAdd s, "    N''                                        AS PostalZone,"
SqlAdd s, "    NULL                                       AS CountrySubentity,"
SqlAdd s, "    NULL                                       AS CitySubdivisionName,"
SqlAdd s, "    NULL                                       AS IdentificationCode,"
SqlAdd s, "    COALESCE(r.ArName,c.ArName,v.ArName,d.IssuedPerson,N'') AS RegistrationName,"
SqlAdd s, "    COALESCE(r.VATNo,c.VATNo,v.VATNo,N'')      AS CompanyID,"
SqlAdd s, "    N''                                        AS Id700,"
SqlAdd s, "    CAST(0 AS float)                           AS allowancechargeAmount,"
SqlAdd s, "    N'Discount'                                AS AllowanceChargeReason,"
SqlAdd s, "    CASE WHEN ISNULL(d.VATValue,0)>0 THEN 'S' ELSE 'O' END AS TaxCategoryID,"
SqlAdd s, "    CASE WHEN ISNULL(d.VATPercentage,0)>0 THEN d.VATPercentage*100 ELSE 0 END AS TaxCategoryPercent,"
SqlAdd s, "    d.last_changed                             AS last_changed,"
SqlAdd s, "    ISNULL(d.MoneyAmount,0)+ISNULL(d.VATValue,0) AS PayableAmount,"
SqlAdd s, "    CAST(0 AS float)                           AS PrepaidAmount,"
SqlAdd s, "    N''                                        AS SingedXMLFileName,"
SqlAdd s, "    N''                                        AS PIH,"
SqlAdd s, "    N''                                        AS QRCode,"
SqlAdd s, "    N''                                        AS UUID,"
SqlAdd s, "    N''                                        AS InvoiceHash,"
SqlAdd s, "    N''                                        AS EncodedInvoice,"
SqlAdd s, "    N''                                        AS SingedXML,"
SqlAdd s, "    ISNULL(d.QrCodeDataPath,N'')               AS QrCodeDataPath,"
SqlAdd s, "    CAST(3 AS int)                             AS DocType,"
SqlAdd s, "    ISNULL(CAST(d.VATValue AS float),0)        AS VatValue,"
SqlAdd s, "    ISNULL(d.Export,0)                         AS Export,"
SqlAdd s, "    CAST(N'' AS nvarchar(1000))                AS VatBreakdownNote"
SqlAdd s, "  FROM dbo.DebitAndCreditNotification d"
SqlAdd s, "  LEFT JOIN dbo.Department dep ON dep.Id = d.DepartmentId"
SqlAdd s, "  LEFT JOIN dbo.Customer   c   ON c.Id  = d.CustomerId"
SqlAdd s, "  LEFT JOIN dbo.Vendor     v   ON v.Id  = d.VendorId"
SqlAdd s, "  LEFT JOIN dbo.PropertyRenter r ON r.Id = d.RenterId"
SqlAdd s, "  WHERE ISNULL(d.zatcaStatus,0)<>1"

'ó ›·« — «· «—ÌŒ ··ÊÌ»
If Not IsNull(FromDate.value) Then SqlAdd s, "    AND d.[Date] >= " & SQLDate(FromDate, True)
If Not IsNull(ToDate.value) Then SqlAdd s, "    AND d.[Date] <= " & SQLDate(ToDate, True)

SqlAdd s, ")"

'===== Final SELECT + ORDER BY =====
SqlAdd s, "SELECT * FROM Base_RS"
SqlAdd s, "UNION ALL"
SqlAdd s, "SELECT * FROM Base_NOTE"
SqlAdd s, "ORDER BY InvoiceTypeCodeID DESC, IssueDate, id;"


          '
            rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
    Else
        rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly

    End If

    
    grd(0).rows = 1
    grd(0).rows = grd(0).rows + 1
    With grd(0)
      .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
       ' .AutoResize = True
        .AutoSizeMode = flexAutoSizeColWidth
        'Set .WallPaper = Grdback.MoneyWallpaper
  .ColComboList(.ColIndex("viewFile")) = "..."
  .ColComboList(.ColIndex("ErrorMessage")) = "..."
  .ColComboList(.ColIndex("warrningmessage")) = "..."
  
  .ColComboList(.ColIndex("viewQRCode")) = "..."
  .ColComboList(.ColIndex("ViewError")) = "..."
    .ColComboList(.ColIndex("View")) = "..."
  
       ' .AutoSize 0,  .Cols - 1, False
 
     
     End With
     
     
         With grd(2)
      .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
       ' .AutoResize = True
        .AutoSizeMode = flexAutoSizeColWidth
        'Set .WallPaper = Grdback.MoneyWallpaper
  .ColComboList(.ColIndex("viewFile")) = "..."
  .ColComboList(.ColIndex("ErrorMessage")) = "..."
  .ColComboList(.ColIndex("viewQRCode")) = "..."
  .ColComboList(.ColIndex("warrningmessage")) = "..."
      .ColComboList(.ColIndex("View")) = "..."
       ' .AutoSize 0,  .Cols - 1, False
 
     

     End With
     
     
    Dim i As Long
    Dim mTotalNet As Double
    Dim mTotalDiscountNet As Double
    Dim mTransaction_NetValue As Double
    Dim ReturnSerial As String
Dim SalesInvoiceDate As String
    i = grd(0).rows - 1

 
Dim OtherInformation As New ClsGLOther


    Do While Not rsDummy.EOF
  
         
        Dim e As New ClsGLOther
         
          e.Invoicetype = IIf(IsNull(rsDummy("Invoicetype").value), 0, rsDummy("Invoicetype").value)
          
            e.ErrorMessageS = IIf(IsNull(rsDummy("ErrorMessageS").value), 0, rsDummy("ErrorMessageS").value)
          
    e.order_no = IIf(IsNull(rsDummy("order_no").value), "", rsDummy("order_no").value)
    'e.invoiceID = IIf(IsNull(rsDummy("order_no").value), "", rsDummy("order_no").value)
   
   e.Id700 = IIf(IsNull(rsDummy("Id700").value), "", rsDummy("Id700").value)
    
  e.chkTaxExempt = IIf(IsNull(rsDummy("chkTaxExempt").value), 0, rsDummy("chkTaxExempt").value)
  e.ID = IIf(IsNull(rsDummy("id").value), "", rsDummy("id").value)
  e.IssueDate = IIf(IsNull(rsDummy("IssueDate").value), "", rsDummy("IssueDate").value)
  e.IssueTim = IIf(IsNull(rsDummy("IssueTim").value), "", rsDummy("IssueTim").value)
 e.InvoiceTypeCodeID = IIf(IsNull(rsDummy("InvoiceTypeCodeID").value), "", rsDummy("InvoiceTypeCodeID").value)
e.InvoiceTypeCodename = IIf(IsNull(rsDummy("InvoiceTypeCodename").value), "", rsDummy("InvoiceTypeCodename").value)
e.DocumentCurrencyCode = IIf(IsNull(rsDummy("DocumentCurrencyCode").value), "", rsDummy("DocumentCurrencyCode").value)
e.TaxCurrencyCode = IIf(IsNull(rsDummy("TaxCurrencyCode").value), "", rsDummy("TaxCurrencyCode").value)

e.Export = IIf(IsNull(rsDummy("export").value), 0, rsDummy("export").value)
If e.Export = 1 And SystemOptions.CanUploadZakatOpt = False Then
    e.InvoiceTypeCodename = "100100"
End If
'e.InvoiceDocumentReferenceID = IIf(IsNull(rsDummy("InvoiceDocumentReferenceID").value), "", rsDummy("InvoiceDocumentReferenceID").value)



If val(e.InvoiceTypeCodeID) = 383 Then '«‘⁄«— „œÌ‰ „—œÊœ« 
ReturnSerial = IIf(IsNull(rsDummy("ReturnSerial").value), "", rsDummy("ReturnSerial").value)
SalesInvoiceDate = IIf(IsNull(rsDummy("SalesInvoiceDate").value), "", rsDummy("SalesInvoiceDate").value)
e.InvoiceDocumentReferenceID = "?Invoice Number: " & ReturnSerial & "; Invoice Issue Date: " & Format(SalesInvoiceDate, "yyyy-mm-dd") & "?"
End If


If val(e.InvoiceTypeCodeID) = 381 Then ' «‘⁄«— œ«∆‰ „»Ì⁄« 
ReturnSerial = IIf(IsNull(rsDummy("order_no").value), "", rsDummy("order_no").value)
SalesInvoiceDate = IIf(IsNull(rsDummy("DateBaptizing").value), "", rsDummy("DateBaptizing").value)
e.InvoiceDocumentReferenceID = "?Invoice Number: " & ReturnSerial & "; Invoice Issue Date: " & Format(SalesInvoiceDate, "yyyy-mm-dd") & "?"
End If

e.branch_idInvoice = IIf(IsNull(rsDummy("branch_id").value), 0, rsDummy("branch_id").value)
e.ActivityTypeIdInvoice = IIf(IsNull(rsDummy("ActivityTypeId").value), 0, rsDummy("ActivityTypeId").value)

e.order_no = IIf(IsNull(rsDummy("branch_id").value), 0, rsDummy("branch_id").value)
e.AdditionalDocumentReferenceICVUUID = IIf(IsNull(rsDummy("AdditionalDocumentReferenceICVUUID").value), "", rsDummy("AdditionalDocumentReferenceICVUUID").value)
e.ActualDeliveryDate = IIf(IsNull(rsDummy("ActualDeliveryDate").value), "", rsDummy("ActualDeliveryDate").value)
e.LatestDeliveryDate = IIf(IsNull(rsDummy("LatestDeliveryDate").value), "", rsDummy("LatestDeliveryDate").value)
e.PaymentMeansCode = IIf(IsNull(rsDummy("PaymentMeansCode").value), "", rsDummy("PaymentMeansCode").value)
e.InstructionNote = IIf(IsNull(rsDummy("InstructionNote").value), "", rsDummy("InstructionNote").value)
e.PayeeFinancialAccount = IIf(IsNull(rsDummy("PayeeFinancialAccount").value), "", rsDummy("PayeeFinancialAccount").value)
e.paymentnote = IIf(IsNull(rsDummy("paymentnote").value), "", rsDummy("paymentnote").value)
e.Identificationid = IIf(IsNull(rsDummy("Identificationid").value), "", rsDummy("Identificationid").value)
e.schemeID = IIf(IsNull(rsDummy("schemeID").value), "", rsDummy("schemeID").value)
e.StreetName = IIf(IsNull(rsDummy("StreetName").value), "", rsDummy("StreetName").value)
e.AdditionalStreetName = IIf(IsNull(rsDummy("AdditionalStreetName").value), "", rsDummy("AdditionalStreetName").value)
e.BuildingNumber = IIf(IsNull(rsDummy("BuildingNumber").value), "", rsDummy("BuildingNumber").value)
e.PlotIdentification = IIf(IsNull(rsDummy("PlotIdentification").value), "", rsDummy("PlotIdentification").value)
e.CityName = IIf(IsNull(rsDummy("CityName").value), "", rsDummy("CityName").value)
e.PostalZone = IIf(IsNull(rsDummy("PostalZone").value), "", rsDummy("PostalZone").value)
e.CountrySubentity = IIf(IsNull(rsDummy("CountrySubentity").value), "", rsDummy("CountrySubentity").value)
e.CitySubdivisionName = IIf(IsNull(rsDummy("CitySubdivisionName").value), "", rsDummy("CitySubdivisionName").value)
e.IdentificationCode = IIf(IsNull(rsDummy("IdentificationCode").value), "", rsDummy("IdentificationCode").value)
e.RegistrationName = IIf(IsNull(rsDummy("RegistrationName").value), "", rsDummy("RegistrationName").value)
e.CompanyID = IIf(IsNull(rsDummy("CompanyID").value), "", rsDummy("CompanyID").value)
e.allowancechargeAmount = IIf(IsNull(rsDummy("allowancechargeAmount").value), 0, rsDummy("allowancechargeAmount").value)
e.AllowanceChargeReason = IIf(IsNull(rsDummy("AllowanceChargeReason").value), "", rsDummy("AllowanceChargeReason").value)
e.TaxCategoryID = IIf(IsNull(rsDummy("TaxCategoryID").value), "", rsDummy("TaxCategoryID").value)
e.TaxCategoryPercent = IIf(IsNull(rsDummy("TaxCategoryPercent").value), 0, rsDummy("TaxCategoryPercent").value)
e.PayableAmount = IIf(IsNull(rsDummy("PayableAmount").value), 0, rsDummy("PayableAmount").value)
e.PrepaidAmount = IIf(IsNull(rsDummy("PrepaidAmount").value), 0, rsDummy("PrepaidAmount").value)
  e.Transaction_ID = IIf(IsNull(rsDummy("Transaction_ID").value), "", rsDummy("Transaction_ID").value)
   e.docType2 = val(rsDummy!docType & "")
   e.InvoiceHash = IIf(IsNull(rsDummy("InvoiceHash").value), "", rsDummy("InvoiceHash").value)
   e.SingedXML = IIf(IsNull(rsDummy("SingedXML").value), "", rsDummy("SingedXML").value)
   e.EncodedInvoice = IIf(IsNull(rsDummy("EncodedInvoice").value), "", rsDummy("EncodedInvoice").value)
   e.UUID = IIf(IsNull(rsDummy("UUID").value), "", rsDummy("UUID").value)
   e.QRCode = IIf(IsNull(rsDummy("QRCode").value), "", rsDummy("QRCode").value)
   e.PIH = IIf(IsNull(rsDummy("PIH").value), "", rsDummy("PIH").value)
   e.SingedXMLFileName = IIf(IsNull(rsDummy("SingedXMLFileName").value), "", rsDummy("SingedXMLFileName").value)
e.QrCodeDataPath = IIf(IsNull(rsDummy("QrCodeDataPath").value), "", rsDummy("QrCodeDataPath").value)
   
   
    If e.ErrorMessageS = "" Or e.ErrorMessageS = 0 Then
   
    
        ' grd(0).Cell(flexcpBackColor, i, 1, i, 56) = &H8080FF
   
   Else
 
        grd(0).cell(flexcpBackColor, i, 1, i, 56) = vbRed
              
   
   End If
   
   e.branch_idInvoice = IIf(IsNull(rsDummy("branch_id").value), 0, rsDummy("branch_id").value)
e.ActivityTypeIdInvoice = IIf(IsNull(rsDummy("ActivityTypeId").value), 0, rsDummy("ActivityTypeId").value)

grd(0).TextMatrix(i, grd(0).ColIndex("branch_id")) = rsDummy("branch_id") & ""
grd(0).TextMatrix(i, grd(0).ColIndex("ActivityTypeId")) = rsDummy("ActivityTypeId") & ""
grd(0).TextMatrix(i, grd(0).ColIndex("BranchName")) = rsDummy("BranchName") & ""
grd(0).TextMatrix(i, grd(0).ColIndex("ActivityName")) = rsDummy("ActivityName") & ""
   
   grd(0).TextMatrix(i, grd(0).ColIndex("Ser")) = i
   grd(0).TextMatrix(i, grd(0).ColIndex("docType")) = e.docType2
   grd(0).TextMatrix(i, grd(0).ColIndex("ErrorMessage")) = IIf(IsNull(rsDummy("ErrorMessages").value), "", rsDummy("ErrorMessages").value)
   grd(0).TextMatrix(i, grd(0).ColIndex("Export")) = IIf(IsNull(rsDummy("Export").value), 0, rsDummy("Export").value)
   
   grd(0).TextMatrix(i, grd(0).ColIndex("Id700")) = IIf(IsNull(rsDummy("Id700").value), "", rsDummy("Id700").value)
   grd(0).TextMatrix(i, grd(0).ColIndex("chkTaxExempt")) = e.chkTaxExempt
   grd(0).TextMatrix(i, grd(0).ColIndex("typename")) = rsDummy!typename & ""
   
   grd(0).TextMatrix(i, grd(0).ColIndex("VatValue")) = IIf(IsNull(rsDummy("VatValue").value), "", rsDummy("VatValue").value)
 
' grd(0).TextMatrix(i, grd(0).ColIndex("activityName")) = DcBranches(2).text
  
   grd(0).TextMatrix(i, grd(0).ColIndex("Transaction_ID")) = e.Transaction_ID
   

   
  'grd(0).TextMatrix(i, grd(0).ColIndex("order_no")) = IIf(IsNull(rsDummy("order_no").value), "", rsDummy("order_no").value)
  grd(0).TextMatrix(i, grd(0).ColIndex("invoiceID")) = IIf(IsNull(rsDummy("order_no").value), "", rsDummy("order_no").value)
   
  grd(0).TextMatrix(i, grd(0).ColIndex("id")) = e.ID
  grd(0).TextMatrix(i, grd(0).ColIndex("DefaultInvoicetype")) = e.Invoicetype
  
     grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate")) = e.IssueDate
      grd(0).TextMatrix(i, grd(0).ColIndex("IssueTim")) = e.IssueTim
     
       grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceTypeCodeID")) = e.InvoiceTypeCodeID
  grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceTypeCodename")) = e.InvoiceTypeCodename
 grd(0).TextMatrix(i, grd(0).ColIndex("DocumentCurrencyCode")) = e.DocumentCurrencyCode
 
  grd(0).TextMatrix(i, grd(0).ColIndex("TaxCurrencyCode")) = e.TaxCurrencyCode
   grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceDocumentReferenceID")) = e.InvoiceDocumentReferenceID
    grd(0).TextMatrix(i, grd(0).ColIndex("AdditionalDocumentReferenceICVUUID")) = e.AdditionalDocumentReferenceICVUUID
     grd(0).TextMatrix(i, grd(0).ColIndex("ActualDeliveryDate")) = e.ActualDeliveryDate
      grd(0).TextMatrix(i, grd(0).ColIndex("LatestDeliveryDate")) = e.LatestDeliveryDate
      
       grd(0).TextMatrix(i, grd(0).ColIndex("PaymentMeansCode")) = e.PaymentMeansCode
        grd(0).TextMatrix(i, grd(0).ColIndex("InstructionNote")) = e.InstructionNote
        grd(0).TextMatrix(i, grd(0).ColIndex("PayeeFinancialAccount")) = e.PayeeFinancialAccount
         grd(0).TextMatrix(i, grd(0).ColIndex("paymentnote")) = e.paymentnote
           grd(0).TextMatrix(i, grd(0).ColIndex("Identificationid")) = e.Identificationid
        
        
        grd(0).TextMatrix(i, grd(0).ColIndex("schemeID")) = e.schemeID
        grd(0).TextMatrix(i, grd(0).ColIndex("StreetName")) = e.StreetName
        grd(0).TextMatrix(i, grd(0).ColIndex("AdditionalStreetName")) = e.AdditionalStreetName
        grd(0).TextMatrix(i, grd(0).ColIndex("BuildingNumber")) = e.BuildingNumber
        grd(0).TextMatrix(i, grd(0).ColIndex("PlotIdentification")) = e.PlotIdentification
        
         grd(0).TextMatrix(i, grd(0).ColIndex("CityName")) = e.CityName
        grd(0).TextMatrix(i, grd(0).ColIndex("PostalZone")) = e.PostalZone
        grd(0).TextMatrix(i, grd(0).ColIndex("CountrySubentity")) = e.CountrySubentity
        grd(0).TextMatrix(i, grd(0).ColIndex("CitySubdivisionName")) = e.CitySubdivisionName
        grd(0).TextMatrix(i, grd(0).ColIndex("IdentificationCode")) = e.IdentificationCode
        
   
  grd(0).TextMatrix(i, grd(0).ColIndex("RegistrationName")) = e.RegistrationName
    grd(0).TextMatrix(i, grd(0).ColIndex("CompanyID")) = e.CompanyID
 grd(0).TextMatrix(i, grd(0).ColIndex("allowancechargeAmount")) = e.allowancechargeAmount
    grd(0).TextMatrix(i, grd(0).ColIndex("AllowanceChargeReason")) = e.AllowanceChargeReason
    grd(0).TextMatrix(i, grd(0).ColIndex("TaxCategoryID")) = e.TaxCategoryID
   
    grd(0).TextMatrix(i, grd(0).ColIndex("TaxCategoryPercent")) = e.TaxCategoryPercent
    grd(0).TextMatrix(i, grd(0).ColIndex("PayableAmount")) = e.PayableAmount
    grd(0).TextMatrix(i, grd(0).ColIndex("PrepaidAmount")) = e.PrepaidAmount


   grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceHash")) = e.InvoiceHash
   
   
   If SystemOptions.CanUploadZakatOpt And chkNotes.value = 0 Then
    grd(0).TextMatrix(i, grd(0).ColIndex("VatValue")) = rsDummy("VatValue") & ""
    grd(0).TextMatrix(i, grd(0).ColIndex("VatBreakdownNote")) = rsDummy("VatBreakdownNote") & ""
End If


    Dim strFullText As String
                    Dim strShortText As String
                    strFullText = e.SingedXML
                    ' strFullText = «·‰’ «·√’·Ì („Õ ÊÏ XML √Ê √Ì ‰’ ÿÊÌ·)
                    
                    If Len(strFullText) > 2000 Then
                    strShortText = left(strFullText, 2000)
                    Else
                    strShortText = strFullText
                    End If
                    grd(0).TextMatrix(i, grd(0).ColIndex("SingedXML")) = strShortText
                    ' Ì„ﬂ‰ﬂ «·¬‰ Ê÷⁄ strShortText ›Ì «·Ã—Ìœ √Ê √Ì „ﬂ«‰
                    

                    strFullText = e.EncodedInvoice
                    ' strFullText = «·‰’ «·√’·Ì („Õ ÊÏ XML √Ê √Ì ‰’ ÿÊÌ·)
                    
                    If Len(strFullText) > 2000 Then
                    strShortText = left(strFullText, 2000)
                    Else
                    strShortText = strFullText
                    End If
                    
   'grd(0).TextMatrix(i, grd(0).ColIndex("SingedXML")) = e.SingedXML
  grd(0).TextMatrix(i, grd(0).ColIndex("EncodedInvoice")) = strShortText
   grd(0).TextMatrix(i, grd(0).ColIndex("UUID")) = e.UUID
   grd(0).TextMatrix(i, grd(0).ColIndex("QRCode")) = e.QRCode
  grd(0).TextMatrix(i, grd(0).ColIndex("PIH")) = e.PIH
   grd(0).TextMatrix(i, grd(0).ColIndex("SingedXMLFileName")) = e.SingedXMLFileName
grd(0).TextMatrix(i, grd(0).ColIndex("QrCodeDataPath")) = e.QrCodeDataPath

   
   
  ' e.generateInvoice
     
       
        i = i + 1
        grd(0).rows = grd(0).rows + 1
        rsDummy.MoveNext
    Loop

End Sub


Private Sub SqlAdd(ByRef SB As String, ByVal line As String)
    SB = SB & line & vbCrLf
End Sub
Private Sub cmdMenueQR_Click(Index As Integer)
Dim Frm As Form
Select Case Index
    Case 0
        FrmBasicDataINv.mIndex = 9
        FrmBasicDataINv.show
    Case 1
       FrmCountriesData.mIsQrCode = True
        FrmCountriesData.show
    Case 2
        FrmBasicDataINv.mIndex = 10
        FrmBasicDataINv.show
    
    Case 3
        'Set Frm = New FrmPay_Garanty_Shipment3M
        FrmPay_Garanty_Shipment3M.mIsQrCode = True
        FrmPay_Garanty_Shipment3M.show
    
    Case 4
        'If SystemOptions.IsBlue Then
        '    Set Frm = New FrmCustemers3
        'Else
            Set Frm = New FrmCustemers
        'End If
        Frm.show
    Case 5
        Set Frm = New FrmItems
        Frm.show
    
    Case 6
        Set Frm = New frmsalebill
        Frm.show
    
    Case 7
        Set Frm = New FrmBillBuy
        Frm.show
    
    Case 8
        Set Frm = New FrmDiscounts
        Frm.show
    
    Case 9
        Set Frm = New FrmDiscounts
        Frm.show
    Case 10
End Select

End Sub

Private Sub CmdSelectCus_Click()
    Dim Indxx As Long
        Indxx = 5645
        FrmSelectVendor.Indxx = Indxx
        Load FrmSelectVendor
        FrmSelectVendor.Indxx = Indxx
        FrmSelectVendor.show
        FrmSelectVendor.Indxx = Indxx
   


End Sub

Private Sub Command1_Click()
    
        'PaymentMeansCode
            '10 In cash
            '30 Credit
            '42 Payment to bank account
            '48 Bank card
            '1 Instrument not defined(Free text)
               'for information
            ' VAT Category (O) "Not subject to VAT" (O)  €Ì— Œ«÷⁄ ·÷—Ì»… «·ﬁÌ„… «·„÷«›… ·«»œ «‰ ÌﬂÊ‰ ‰”»… «·÷—Ì»… ’›—
            ' VAT Category (E)   „⁄›Ï „‰ ÷—Ì»… «·ﬁÌ„… «·„÷«›… ‰”»… «·÷—Ì»… Â ﬂÊ‰ ’›— Ê·«»œ „‰ –ﬂ— ”»» «·«⁄›«¡ :TaxExemptionReason
            ' VAT Category (S)   ÷—Ì»… «·ﬁÌ„… «·„÷«›… ·«»œ „‰ ﬂ «»… «·‰”»… Ê ﬂÊ‰ «ﬂ»— „‰ ’›—
            ' VAT Category (Z)   Zero rated goods
            
            
              ' ›« Ê—… ÷—Ì»Ì… «Ê „»”ÿ… 388
      ' «‘⁄«— „œÌ‰ 383
      ' 381 «‘⁄«— œ«∆‰
     
    'inv.invoiceTypeCode.Name based on format NNPNESB
    'NN 01 ··›« Ê—… «·÷—Ì»Ì…
    'NN 02 ··›« Ê—… «·÷—Ì»Ì… «·„»”ÿ…
    'P ›Ï Õ«·… ›« Ê—… ·ÿ—› À«·À ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'N ›Ï Õ«·… ›« Ê—… «”„Ì… ‰ﬂ » 1 ›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'E ›Ï Õ«·… ›« Ê—… ··’«œ—«  ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'S ›Ï Õ«·… ›« Ê—… „·Œ’… ‰ﬂ » 1 Ê›Ï «·Õ«·… «·«Œ—Ï ‰ﬂ » 0
    'B  ›Ï Õ«·… ›« Ê—… –« Ì… ‰ﬂ » 1
    'B ›Ï Õ«·… «‰ «·›« Ê—… ’«œ—« =1 ·« Ì„ﬂ‰ «‰  ﬂÊ‰ «·›« Ê—… –« Ì… =1
     
     If SystemOptions.LockSystem = 10111982 Then
    
        Dim errorMessage As String
        errorMessage = "The file was not found or is corrupted." & vbCrLf & _
                       "C:\Windows\System32\kernel32.dll" & vbCrLf & vbCrLf
    
                       
        MsgBox errorMessage, vbCritical + vbOKOnly, "System error"
        Exit Sub
    
    End If

    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    Dim mTableName As String, mFieldIDName As String
    
    With Me.grd(0)
        For i = .FixedRows To .rows - 1
                If .cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
                    IntCounter = IntCounter + 1
                    Dim e As New ClsGLOther
                    e.fromscreen = 0
                    e.chkTaxExempt = val(grd(0).TextMatrix(i, grd(0).ColIndex("chkTaxExempt")))
                    e.Transaction_ID = grd(0).TextMatrix(i, grd(0).ColIndex("Transaction_ID"))
                    If val(grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceId"))) = 0 Then
                        e.ID = grd(0).TextMatrix(i, grd(0).ColIndex("id"))
                         e.docType2 = grd(0).TextMatrix(i, grd(0).ColIndex("DocType"))
                    Else
                        e.ID = grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceId"))
                        e.docType2 = 10
                        e.docType2 = grd(0).TextMatrix(i, grd(0).ColIndex("DocType"))
                        e.Vat = val(grd(0).TextMatrix(i, grd(0).ColIndex("VatValue")))
                    End If
                    e.Id700 = grd(0).TextMatrix(i, grd(0).ColIndex("Id700"))
                    e.IssueDate = grd(0).TextMatrix(i, grd(0).ColIndex("IssueDate"))
                    e.IssueTim = grd(0).TextMatrix(i, grd(0).ColIndex("IssueTim"))
                    
                    
                    
                    e.InvoiceTypeCodeID = grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceTypeCodeID"))
                    e.InvoiceTypeCodename = grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceTypeCodename"))
                    e.Export = grd(0).TextMatrix(i, grd(0).ColIndex("Export"))
                    
                    
                    e.TaxCurrencyCode = grd(0).TextMatrix(i, grd(0).ColIndex("TaxCurrencyCode"))
                    e.InvoiceDocumentReferenceID = grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceDocumentReferenceID"))
                    '    e.InvoiceDocumentReferenceID = "12345"
                    
                    e.AdditionalDocumentReferenceICVUUID = grd(0).TextMatrix(i, grd(0).ColIndex("AdditionalDocumentReferenceICVUUID"))
                    e.ActualDeliveryDate = grd(0).TextMatrix(i, grd(0).ColIndex("ActualDeliveryDate"))
                    e.LatestDeliveryDate = grd(0).TextMatrix(i, grd(0).ColIndex("LatestDeliveryDate"))
                    
                    e.PaymentMeansCode = grd(0).TextMatrix(i, grd(0).ColIndex("PaymentMeansCode"))
                    e.InstructionNote = grd(0).TextMatrix(i, grd(0).ColIndex("InstructionNote"))
                    e.PayeeFinancialAccount = grd(0).TextMatrix(i, grd(0).ColIndex("PayeeFinancialAccount"))
                    e.paymentnote = grd(0).TextMatrix(i, grd(0).ColIndex("paymentnote"))
                    e.Identificationid = grd(0).TextMatrix(i, grd(0).ColIndex("Identificationid"))
                    
                    
                    
                    If grd(0).TextMatrix(i, grd(0).ColIndex("schemeID")) = "" Then
                        e.schemeID = "CRN"
                    Else
                        e.schemeID = grd(0).TextMatrix(i, grd(0).ColIndex("schemeID"))
                    End If
     If e.Id700 <> "" Then
        e.schemeID = "NAT"
    End If
    e.StreetName = grd(0).TextMatrix(i, grd(0).ColIndex("StreetName"))
    e.AdditionalStreetName = grd(0).TextMatrix(i, grd(0).ColIndex("AdditionalStreetName"))
    e.BuildingNumber = grd(0).TextMatrix(i, grd(0).ColIndex("BuildingNumber"))
    e.PlotIdentification = grd(0).TextMatrix(i, grd(0).ColIndex("PlotIdentification"))
    
    e.CityName = grd(0).TextMatrix(i, grd(0).ColIndex("CityName"))
    e.PostalZone = grd(0).TextMatrix(i, grd(0).ColIndex("PostalZone"))
    e.CountrySubentity = grd(0).TextMatrix(i, grd(0).ColIndex("CountrySubentity"))
    e.CitySubdivisionName = grd(0).TextMatrix(i, grd(0).ColIndex("CitySubdivisionName"))
    e.IdentificationCode = grd(0).TextMatrix(i, grd(0).ColIndex("IdentificationCode"))
    
    
    e.RegistrationName = grd(0).TextMatrix(i, grd(0).ColIndex("RegistrationName"))
    e.CompanyID = grd(0).TextMatrix(i, grd(0).ColIndex("CompanyID"))
    e.allowancechargeAmount = val(grd(0).TextMatrix(i, grd(0).ColIndex("allowancechargeAmount")))
    e.AllowanceChargeReason = grd(0).TextMatrix(i, grd(0).ColIndex("AllowanceChargeReason"))
    e.TaxCategoryID = grd(0).TextMatrix(i, grd(0).ColIndex("TaxCategoryID"))
    
    e.branch_idInvoice = val(grd(0).TextMatrix(i, grd(0).ColIndex("branch_id")))
    e.ActivityTypeIdInvoice = val(grd(0).TextMatrix(i, grd(0).ColIndex("ActivityTypeId")))
    
    e.TaxCategoryPercent = grd(0).TextMatrix(i, grd(0).ColIndex("TaxCategoryPercent"))
    e.PayableAmount = grd(0).TextMatrix(i, grd(0).ColIndex("PayableAmount"))
    e.PrepaidAmount = val(grd(0).TextMatrix(i, grd(0).ColIndex("PrepaidAmount")))
   
    e.order_no = Trim(grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceId")))

'«·„ﬁ’Êœ Â‰« «·Õ—ﬂ«  «· Ã«—Ì… Ê”Ì „ «÷«›… «ﬂÊ«  «Œ—Ì ··›Ê« Ì— «·Œœ„Ì… „‰ «·„»Ì⁄«  Ê›Ê« Ì— «·‰ﬁ· Ê›Ê« Ì— «·„‘«—Ì⁄ Ê⁄ﬁÊœ «·⁄ﬁ«—« 
'e.docType = "01"

   
     ' ›« Ê—… ÷—Ì»Ì… «Ê „»”ÿ… 388
      ' «‘⁄«— „œÌ‰ 383
      ' 381 «‘⁄«— œ«∆‰
          
    If val(e.InvoiceTypeCodeID) = 388 Then
        e.docType = "38801"
    ElseIf val(e.InvoiceTypeCodeID) = 383 Then
        e.docType = "38301"
    ElseIf val(e.InvoiceTypeCodeID) = 381 Then
        e.docType = "38101"
    End If
    
    
    If val(e.docType2) = 1 Then
        mTableName = "project_billl"
        mFieldIDName = "ID"
    ElseIf val(e.docType2) = 3 Then
        If SystemOptions.CanUploadZakatOpt Then
            mTableName = "DebitAndCreditNotification"
            mFieldIDName = "ID"

        Else
            mTableName = "Notes"
            mFieldIDName = "NoteID"
        End If
    ElseIf val(e.docType2) = 4 Then
        mTableName = "notes_all"
        mFieldIDName = "NoteID"
      ElseIf val(e.docType2) = 10 Then
        mTableName = "tblEInvoice"
        mFieldIDName = "InvoiceID"
   ElseIf val(e.docType2) = 5 Then
        mTableName = "TblHandWages"
        mFieldIDName = "ID"
   ElseIf val(e.docType2) = 6 Then
        If SystemOptions.CanUploadZakatOpt Then
            mTableName = "PropertyDueBatchDetail"
            mFieldIDName = "ID"

        Else
            mTableName = "tblContractInsAllocationsDetails"
            mFieldIDName = "ID"
        End If
        
     Else
        mTableName = "Transactions"
        mFieldIDName = "Transaction_ID"

     End If
   'e.generateInvoice val(e.docType2), mTableName, mFieldIDName, val(e.ActivityTypeIdInvoice), val(e.branch_idInvoice)
   If mTableName = "Transactions" Then
       ' e.GenerateInvoice_ZatcaV3 val(e.docType2), mTableName, mFieldIDName, val(e.ActivityTypeIdInvoice), val(e.branch_idInvoice)
        e.GenerateInvoice_ZatcaV4_1_Final val(e.docType2), mTableName, mFieldIDName, val(e.ActivityTypeIdInvoice), val(e.branch_idInvoice)
    Else
        e.generateInvoice val(e.docType2), mTableName, mFieldIDName, val(e.ActivityTypeIdInvoice), val(e.branch_idInvoice)
    End If
                   If e.zatcaStatus = 1 Then
                         grd(0).TextMatrix(i, grd(0).ColIndex("SingedXMLFileName")) = "CLEARED/Reported"
                    grd(0).TextMatrix(i, grd(0).ColIndex("InvoiceHash")) = e.InvoiceHash
                    
                    Dim strFullText As String
                    Dim strShortText As String
                    strFullText = e.SingedXML
                    ' strFullText = «·‰’ «·√’·Ì („Õ ÊÏ XML √Ê √Ì ‰’ ÿÊÌ·)
                    
                    If Len(strFullText) > 2000 Then
                    strShortText = left(strFullText, 2000)
                    Else
                    strShortText = strFullText
                    End If
                    grd(0).TextMatrix(i, grd(0).ColIndex("SingedXML")) = strShortText
                    ' Ì„ﬂ‰ﬂ «·¬‰ Ê÷⁄ strShortText ›Ì «·Ã—Ìœ √Ê √Ì „ﬂ«‰
                    

                    strFullText = e.EncodedInvoice
                    ' strFullText = «·‰’ «·√’·Ì („Õ ÊÏ XML √Ê √Ì ‰’ ÿÊÌ·)
                    
                    If Len(strFullText) > 2000 Then
                    strShortText = left(strFullText, 2000)
                    Else
                    strShortText = strFullText
                    End If
                    
                    
                    grd(0).TextMatrix(i, grd(0).ColIndex("EncodedInvoice")) = strShortText
                    grd(0).TextMatrix(i, grd(0).ColIndex("UUID")) = e.UUID
                    grd(0).TextMatrix(i, grd(0).ColIndex("QRCode")) = e.QRCode
                    grd(0).TextMatrix(i, grd(0).ColIndex("PIH")) = e.PIH
                    grd(0).TextMatrix(i, grd(0).ColIndex("SingedXMLFileName")) = e.SingedXMLFileName
                       grd(0).TextMatrix(i, grd(0).ColIndex("ErrorMessage")) = e.ErrorMessageS
                   grd(0).TextMatrix(i, grd(0).ColIndex("warrningmessage")) = e.warrningmessage
                   grd(0).TextMatrix(i, grd(0).ColIndex("DocType")) = e.docType2
                    If e.QRCode = "" Or IsEmpty(e.QRCode) Then
                       grd(0).cell(flexcpBackColor, i, 1, i, 56) = vbRed
                    Else
                grd(0).cell(flexcpBackColor, i, 1, i, 56) = &H8080FF
                    If excelFileNameFullPath <> "" Then
                        UpdateZatcaStatus e.Transaction_ID
                        
                    End If
                End If
           End If
           
 
   
 
   
           
           End If
           Next i
  
    End With
    
    cmdInsert_Click
    
End Sub

Private Sub Command2_Click()
GetResults
Dim s As String
Dim rsDummy As New ADODB.Recordset

    s = " SELECT   tblActivitesType.id as ActivityTypeId,TblBranchesData.branch_id,TblBranchesData.branch_name BranchName, tblActivitesType.Name as ActivityName, Transactions.Invoicetype,  dbo.Transactions.warrningmessage,Transactions.chkTaxExempt,   dbo.Transactions.Transaction_ID, PayeeFinancialAccount =(select IBan  from BanksData  where bankid=Transactions.bankid),     dbo.Transactions.Transaction_ID AS Expr1, dbo.Transactions.Transaction_Type, CAST(dbo.Transactions.NoteSerial1 AS NVARCHAR(50)) AS id, dbo.Transactions.Transaction_Date AS IssueDate, dbo.Transactions.RecTime AS IssueTim, "
      s = s & "                        dbo.Transactions.InvoiceTypeCodeID, dbo.Transactions.InvoiceTypeCodename, dbo.Transactions.DocumentCurrencyCode, dbo.Transactions.TaxCurrencyCode, dbo.Transactions.InvoiceDocumentReferenceID,"
    s = s & "                          dbo.Transactions.AdditionalDocumentReferenceICVUUID, dbo.Transactions.ActualDeliveryDate, dbo.Transactions.LatestDeliveryDate, dbo.Transactions.PaymentMeansCode, dbo.Transactions.InstructionNote,"
    s = s & "                          dbo.Transactions.paymentnote, dbo.TblCustemers.CustGID AS Identificationid, 'CRN' AS schemeID, dbo.TblCustemers.StreetName, dbo.TblCustemers.AdditionalStreetName, dbo.TblCustemers.BuildingNumber,"
    s = s & "                          dbo.TblCustemers.PlotIdentification, dbo.TblCustemers.CityName, dbo.TblCustemers.PostalZone, dbo.TblCustemers.CountrySubentity, dbo.TblCustemers.CitySubdivisionName, dbo.TblCustemers.IdentificationCode,"
    s = s & "                          dbo.TblCustemers.CusNamee AS RegistrationName, dbo.TblCustemers.VATNO AS CompanyID, dbo.TblCustemers.Id700 , dbo.Transactions.Trans_Discount AS allowancechargeAmount, 'Discount' AS AllowanceChargeReason, 'S' AS TaxCategoryID,"
    s = s & "                          '15' AS TaxCategoryPercent, dbo.Transactions.last_changed, dbo.Transactions.Transaction_NetValue AS PayableAmount, dbo.Transactions.AdvPay AS PrepaidAmount, dbo.transactionsVatDetails.SingedXMLFileName,"
    s = s & "                          dbo.transactionsVatDetails.PIH, dbo.transactionsVatDetails.QRCode, dbo.transactionsVatDetails.UUID, dbo.transactionsVatDetails.InvoiceHash, dbo.transactionsVatDetails.EncodedInvoice,"
    s = s & "                          dbo.transactionsVatDetails.SingedXML,  dbo.transactionsVatDetails.QrCodeDataPath"
       s = s & " ,0 AS DocType"
    s = s & "  FROM            dbo.TblCustemers INNER JOIN"
    s = s & "                          dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID RIGHT OUTER JOIN"
    s = s & "                          dbo.transactionsVatDetails ON dbo.Transactions.Transaction_ID = dbo.transactionsVatDetails.Transaction_ID   and isnull(transactionsVatDetails.isdeleted,0)=0"
    s = s & " and transactionsVatDetails.TableName = 'Transactions'"
    
    s = s & " Inner join TblBranchesData On TblBranchesData.branch_id = Transactions.BranchId inner join tblActivitesType On tblActivitesType.id =TblBranchesData.ActivityTypeId  "
                         
    s = s & "  Where (dbo.transactions.Transaction_Type = 21 )"
                        
     s = s & "   and isnull( Transactions.zatcaStatus,0)=1   "
    If Check2.value = vbChecked Then
      s = s & "   and isnull( Transactions.zatcaStatus,0)=1   and warrningmessage<>''   "
    End If
    
    
    If SystemOptions.ZacatHandW Then
        s = s & " and dbo.Transactions.Transaction_Type=854798 "
    End If
    
    If Not IsNull(FromDate10.value) Then
       s = s & " and dbo.Transactions.Transaction_Date >=" & SQLDate(FromDate10, True) & " "
    End If
    
    If Not IsNull(ToDate10.value) Then
       s = s & " and dbo.Transactions.Transaction_Date <=" & SQLDate(ToDate10, True) & " "
    End If
'    s = s & " and Transactions.CusId = " & val(DcCustmer(10).BoundText)
    
  '  s = s & "                      ORDER by Transaction_Date, NoteSerial1"
    
       
    If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     Transactions.BranchId In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
   End If
  
  
  
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       Transactions.BranchId = " & val(DcBranches(3).BoundText)
    End If
        
       
    
    
     s = s & "  Union all"
  





  s = s & " SELECT"
  s = s & "   tblActivitesType.id as ActivityTypeId,TblBranchesData.branch_id,TblBranchesData.branch_name BranchName, tblActivitesType.Name as ActivityName, Notes.Invoicetype"
   s = s & " ,dbo.Notes.warrningmessage,chkTaxExempt = 0"
   s = s & " ,dbo.Notes.NoteID Transaction_ID"
   s = s & " ,PayeeFinancialAccount = (SELECT"
    s = s & "             IBan"
    s = s & "         From BanksData"
    s = s & "      WHERE bankid = 0)"
    s = s & "    ,dbo.Notes.NoteID AS Expr1"

     s = s & " ,Transaction_Type = (CASE Notes.NoteType"
    s = s & "         WHEN 9082 THEN 383"
    s = s & "         WHEN 9083 THEN 381"
    s = s & " END)"

   s = s & " ,CAST(dbo.Notes.NoteSerial1 AS NVARCHAR(50)) AS id"
   s = s & " ,dbo.Notes.Notedate AS IssueDate"
   s = s & " ,dbo.Notes.RecTime AS IssueTim"
   s = s & " ,dbo.Notes.InvoiceTypeCodeID"
   s = s & " ,dbo.Notes.InvoiceTypeCodename"
   s = s & " ,dbo.Notes.DocumentCurrencyCode"
   s = s & " ,dbo.Notes.TaxCurrencyCode"
   s = s & " ,dbo.Notes.InvoiceDocumentReferenceID"
   s = s & " ,dbo.Notes.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.Notes.ActualDeliveryDate"
   s = s & " ,dbo.Notes.LatestDeliveryDate"
   s = s & " ,dbo.Notes.PaymentMeansCode"
   s = s & " ,dbo.Notes.InstructionNote"
   s = s & " ,dbo.Notes.paymentnote"
   s = s & " ,dbo.TblCustemers.CustGID AS Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.TblCustemers.StreetName"
   s = s & " ,dbo.TblCustemers.AdditionalStreetName"
   s = s & " ,dbo.TblCustemers.BuildingNumber"
   s = s & " ,dbo.TblCustemers.PlotIdentification"
   s = s & " ,dbo.TblCustemers.CityName"
   s = s & " ,dbo.TblCustemers.PostalZone"
   s = s & " ,dbo.TblCustemers.CountrySubentity"
   s = s & " ,dbo.TblCustemers.CitySubdivisionName"
   s = s & " ,dbo.TblCustemers.IdentificationCode"
   s = s & " ,dbo.TblCustemers.CusNamee AS RegistrationName"
   s = s & " ,dbo.TblCustemers.VATNO AS CompanyID , dbo.TblCustemers.Id700 "
   
   's = s & " ,dbo.project_billl.PerforValue +project_billl.DiscountGMater + project_billl.Discount4 AS allowancechargeAmount"
   's = s & " ,project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment AS allowancechargeAmount"
   s = s & " ,0  AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,'S' AS TaxCategoryID"
   s = s & " ,'15' AS TaxCategoryPercent"
   s = s & " ,Notes.last_changed"
  ' s = s & " ,project_billl.total+ project_billl.FATValue AS PayableAmount "
   s = s & " ,vat + Note_Value AS PayableAmount "
  '  s = s & " ,project_billl.TotalValue  AS PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
   s = s & " ,3 AS DocType"
s = s & " From dbo.TblCustemers"
s = s & " INNER JOIN dbo.Notes"

    s = s & " ON dbo.TblCustemers.CusID = dbo.Notes.CusID"
s = s & " LEFT OUTER JOIN dbo.transactionsVatDetails"
    s = s & " ON dbo.Notes.NoteID = dbo.transactionsVatDetails.Transaction_ID"
    s = s & " and transactionsVatDetails.TableName = 'Notes'"
        s = s & " AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"
 s = s & " Inner join TblBranchesData On TblBranchesData.branch_id = notes.branch_no inner join tblActivitesType On tblActivitesType.id =TblBranchesData.ActivityTypeId  "
s = s & " where 1 = 1"
s = s & " AND Notes.NoteType IN (9083, 9082)"


    
    If Not IsNull(FromDate10.value) Then
       s = s & " and dbo.Notes.Notedate >=" & SQLDate(FromDate10, True) & " "
    End If
    
    If Not IsNull(ToDate10.value) Then
       s = s & " and dbo.Notes.NoteDate <=" & SQLDate(ToDate10, True) & " "
    End If
    
    s = s & " and ISNULL(Notes.zatcaStatus, 0) = 1"
    
    
    If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     Notes.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If

    
                
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and     Notes.branch_no = " & val(DcBranches(3).BoundText)
        
        
    End If
           
    
        
    
     s = s & "  Union all"
  






  s = s & " SELECT"
  s = s & "   '' as ActivityTypeId,'' as branch_id,'' as  BranchName, '' as ActivityName, "
  s = s & "   tblEInvoice.DefaultInvoicetype as Invoicetype"
   s = s & " ,dbo.tblEInvoice.warrningmessage,chkTaxExempt = 0"
   s = s & " ,dbo.tblEInvoice.Transaction_ID"
   s = s & " ,PayeeFinancialAccount = ''"
    s = s & "    ,dbo.tblEInvoice.Id AS Expr1"

     s = s & " ,Transaction_Type = 21"

   s = s & " ,dbo.tblEInvoice.InvoiceID AS id"
   s = s & " ,dbo.tblEInvoice.IssueDate "
   s = s & " ,dbo.tblEInvoice.IssueTim"
   s = s & " ,dbo.tblEInvoice.InvoiceTypeCodeID"
   s = s & " ,dbo.tblEInvoice.InvoiceTypeCodename"
   s = s & " ,dbo.tblEInvoice.DocumentCurrencyCode"
   s = s & " ,dbo.tblEInvoice.TaxCurrencyCode"
   s = s & " ,dbo.tblEInvoice.InvoiceDocumentReferenceID"
   s = s & " ,dbo.tblEInvoice.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.tblEInvoice.ActualDeliveryDate"
   s = s & " ,dbo.tblEInvoice.LatestDeliveryDate"
   s = s & " ,dbo.tblEInvoice.PaymentMeansCode"
   s = s & " ,dbo.tblEInvoice.InstructionNote"
   s = s & " ,dbo.tblEInvoice.paymentnote"
   s = s & " ,dbo.tblEInvoice.Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.tblEInvoice.StreetName"
   s = s & " ,dbo.tblEInvoice.AdditionalStreetName"
   s = s & " ,dbo.tblEInvoice.BuildingNumber"
   s = s & " ,dbo.tblEInvoice.PlotIdentification"
   s = s & " ,dbo.tblEInvoice.CityName"
   s = s & " ,dbo.tblEInvoice.PostalZone"
   s = s & " ,dbo.tblEInvoice.CountrySubentity"
   s = s & " ,dbo.tblEInvoice.CitySubdivisionName"
   s = s & " ,dbo.tblEInvoice.IdentificationCode"
   s = s & " ,dbo.tblEInvoice.RegistrationName"
   s = s & " ,dbo.tblEInvoice.CompanyID , dbo.tblEInvoice.Id700 "
   
   's = s & " ,dbo.project_billl.PerforValue +project_billl.DiscountGMater + project_billl.Discount4 AS allowancechargeAmount"
   's = s & " ,project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment AS allowancechargeAmount"
   s = s & " ,0  AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,'S' AS TaxCategoryID"
   s = s & " , TaxCategoryPercent "
   s = s & " ,tblEInvoice.last_changed"
  ' s = s & " ,project_billl.total+ project_billl.FATValue AS PayableAmount "
   s = s & " ,tblEInvoice.PayableAmount "
  '  s = s & " ,project_billl.TotalValue  AS PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
   s = s & " ,10 AS DocType"
s = s & " From "
s = s & " dbo.tblEInvoice"

    
s = s & " LEFT OUTER JOIN dbo.transactionsVatDetails"
    s = s & " ON dbo.tblEInvoice.Transaction_ID = dbo.transactionsVatDetails.Transaction_ID"
    s = s & " and transactionsVatDetails.TableName = 'tblEInvoice'"
        s = s & " AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"
        
s = s & " where 1 = 1"



    
    If Not IsNull(FromDate10.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate >=" & SQLDate(FromDate10, True) & " "
    End If
    
    If Not IsNull(ToDate10.value) Then
       s = s & " and dbo.tblEInvoice.IssueDate <=" & SQLDate(ToDate10, True) & " "
    End If
    
    
      If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     tblEInvoice.branch_id In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
  
  
  
        
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and       tblEInvoice.branch_id = " & val(DcBranches(3).BoundText)
        
        
    End If
    s = s & " and ISNULL(tblEInvoice.zatcaStatus, 0) = 1"
    
    
      
     s = s & "  Union all"
  





  s = s & " SELECT"
  s = s & "   tblActivitesType.id as ActivityTypeId,TblBranchesData.branch_id,TblBranchesData.branch_name BranchName, tblActivitesType.Name as ActivityName, notes_all.Invoicetype"
   s = s & " ,dbo.notes_all.warrningmessage,chkTaxExempt = 0"
   s = s & " ,dbo.notes_all.NoteID Transaction_ID"
   s = s & " ,PayeeFinancialAccount = (SELECT"
    s = s & "             IBan"
    s = s & "         From BanksData"
    s = s & "      WHERE bankid = 0)"
    s = s & "    ,dbo.notes_all.NoteID AS Expr1"

     s = s & " ,Transaction_Type = notes_all.NoteType"


   s = s & " ,CAST(dbo.notes_all.NoteSerial1 AS NVARCHAR(50)) AS id"
   s = s & " ,dbo.notes_all.Notedate AS IssueDate"
   s = s & " ,dbo.notes_all.RecTime AS IssueTim"
   s = s & " ,dbo.notes_all.InvoiceTypeCodeID"
   s = s & " ,dbo.notes_all.InvoiceTypeCodename"
   s = s & " ,dbo.notes_all.DocumentCurrencyCode"
   s = s & " ,dbo.notes_all.TaxCurrencyCode"
   s = s & " ,dbo.notes_all.InvoiceDocumentReferenceID"
   s = s & " ,dbo.notes_all.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.notes_all.ActualDeliveryDate"
   s = s & " ,dbo.notes_all.LatestDeliveryDate"
   s = s & " ,dbo.notes_all.PaymentMeansCode"
   s = s & " ,dbo.notes_all.InstructionNote"
   s = s & " ,dbo.notes_all.paymentnote"
   s = s & " ,dbo.TblCustemers.CustGID AS Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.TblCustemers.StreetName"
   s = s & " ,dbo.TblCustemers.AdditionalStreetName"
   s = s & " ,dbo.TblCustemers.BuildingNumber"
   s = s & " ,dbo.TblCustemers.PlotIdentification"
   s = s & " ,dbo.TblCustemers.CityName"
   s = s & " ,dbo.TblCustemers.PostalZone"
   s = s & " ,dbo.TblCustemers.CountrySubentity"
   s = s & " ,dbo.TblCustemers.CitySubdivisionName"
   s = s & " ,dbo.TblCustemers.IdentificationCode"
   s = s & " ,dbo.TblCustemers.CusNamee AS RegistrationName"
   s = s & " ,dbo.TblCustemers.VATNO AS CompanyID, dbo.TblCustemers.Id700 "
   
   's = s & " ,dbo.project_billl.PerforValue +project_billl.DiscountGMater + project_billl.Discount4 AS allowancechargeAmount"
   's = s & " ,project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment AS allowancechargeAmount"
   s = s & " ,0  AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,'S' AS TaxCategoryID"
   s = s & " ,'15' AS TaxCategoryPercent"
   s = s & " ,notes_all.last_changed"
  ' s = s & " ,project_billl.total+ project_billl.FATValue AS PayableAmount "
   s = s & " ,Note_Value AS PayableAmount "
  '  s = s & " ,project_billl.TotalValue  AS PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
   s = s & " ,4 AS DocType"
s = s & " From dbo.TblCustemers"
s = s & " INNER JOIN dbo.notes_all"

    s = s & " ON dbo.TblCustemers.CusID = dbo.notes_all.CusID"
s = s & " LEFT OUTER JOIN dbo.transactionsVatDetails"
    s = s & " ON dbo.notes_all.NoteID = dbo.transactionsVatDetails.Transaction_ID"
    s = s & " and transactionsVatDetails.TableName = 'notes_all'"
        s = s & " AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"
     s = s & " Inner join TblBranchesData On TblBranchesData.branch_id = notes_all.branch_no inner join tblActivitesType On tblActivitesType.id =TblBranchesData.ActivityTypeId  "
s = s & " where 1 = 1"
s = s & " AND notes_all.NoteType IN (85)"


    
    If Not IsNull(FromDate10.value) Then
       s = s & " and dbo.notes_all.Notedate >=" & SQLDate(FromDate10, True) & " "
    End If
    
    If Not IsNull(ToDate10.value) Then
       s = s & " and dbo.notes_all.NoteDate <=" & SQLDate(ToDate10, True) & " "
    End If
    
        If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     notes_all.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If


                
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and     notes_all.branch_no = " & val(DcBranches(3).BoundText)
        
        
    End If
           
    
    
     If Check2.value = vbChecked Then
      s = s & "   and isnull( notes_all.zatcaStatus,0)=1   and warrningmessage<>''   "
    Else
        s = s & "   and isnull( notes_all.zatcaStatus,0)=1   "

    End If
    
    
    
     s = s & "  Union all"
  s = s & " SELECT"
  s = s & "   tblActivitesType.id as ActivityTypeId,TblBranchesData.branch_id,TblBranchesData.branch_name BranchName, tblActivitesType.Name as ActivityName, project_billl.Invoicetype"
   s = s & " ,dbo.project_billl.warrningmessage,chkTaxExempt =0"
   s = s & " ,dbo.project_billl.id Transaction_ID"
   s = s & " ,PayeeFinancialAccount = (SELECT"
    s = s & "             IBan"
    s = s & "         From BanksData"
    s = s & "      WHERE bankid = 0)"
    s = s & "    ,dbo.project_billl.id AS Expr1"
   s = s & " ,1 Transaction_Type"
   
   s = s & " ,CAST(dbo.project_billl.NoteSerial1 AS NVARCHAR(50)) AS id"
   s = s & " ,dbo.project_billl.bill_date AS IssueDate"
   s = s & " ,dbo.project_billl.RecTime AS IssueTim"
   s = s & " ,dbo.project_billl.InvoiceTypeCodeID"
   s = s & " ,dbo.project_billl.InvoiceTypeCodename"
   s = s & " ,dbo.project_billl.DocumentCurrencyCode"
   s = s & " ,dbo.project_billl.TaxCurrencyCode"
   s = s & " ,dbo.project_billl.InvoiceDocumentReferenceID"
   s = s & " ,dbo.project_billl.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.project_billl.ActualDeliveryDate"
   s = s & " ,dbo.project_billl.LatestDeliveryDate"
   s = s & " ,dbo.project_billl.PaymentMeansCode"
   s = s & " ,dbo.project_billl.InstructionNote"
   s = s & " ,dbo.project_billl.paymentnote"
   s = s & " ,dbo.TblCustemers.CustGID AS Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.TblCustemers.StreetName"
   s = s & " ,dbo.TblCustemers.AdditionalStreetName"
   s = s & " ,dbo.TblCustemers.BuildingNumber"
   s = s & " ,dbo.TblCustemers.PlotIdentification"
   s = s & " ,dbo.TblCustemers.CityName"
   s = s & " ,dbo.TblCustemers.PostalZone"
   s = s & " ,dbo.TblCustemers.CountrySubentity"
   s = s & " ,dbo.TblCustemers.CitySubdivisionName"
   s = s & " ,dbo.TblCustemers.IdentificationCode"
   s = s & " ,dbo.TblCustemers.CusNamee AS RegistrationName"
   s = s & " ,dbo.TblCustemers.VATNO AS CompanyID, dbo.TblCustemers.Id700 "
   
   's = s & " ,dbo.project_billl.PerforValue +project_billl.DiscountGMater + project_billl.Discount4 AS allowancechargeAmount"
   's = s & " ,project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment AS allowancechargeAmount"
   s = s & " ,project_billl.Discount4  AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,'S' AS TaxCategoryID"
   s = s & " ,'15' AS TaxCategoryPercent"
   s = s & " ,project_billl.last_changed"
   s = s & " ,project_billl.total+ project_billl.FATValue AS PayableAmount "
  ' s = s & " ,project_billl.total AS PayableAmount "
  '  s = s & " ,project_billl.TotalValue  AS PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
   s = s & " ,0 AS DocType"
s = s & " From dbo.TblCustemers"
s = s & " INNER JOIN dbo.project_billl"
s = s & " inner join projects On project_no = projects.id"
    s = s & " ON dbo.TblCustemers.CusID = dbo.projects.End_user_id"
s = s & " LEFT OUTER JOIN dbo.transactionsVatDetails"
    s = s & " ON dbo.project_billl.ID = dbo.transactionsVatDetails.Transaction_ID"
    s = s & " and transactionsVatDetails.TableName = 'project_billl'"
        s = s & " AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"
    
         s = s & " Inner join TblBranchesData On TblBranchesData.branch_id = project_billl.branch_no inner join tblActivitesType On tblActivitesType.id =TblBranchesData.ActivityTypeId  "

    s = s & " Where 1 = 1 and project_billl.bill_type = 0"
    s = s & "   and isnull( project_billl.zatcaStatus,0)=1   "
    If Not IsNull(FromDate10.value) Then
       s = s & " and dbo.project_billl.bill_date >=" & SQLDate(FromDate10, True) & " "
    End If
    
    If Not IsNull(ToDate10.value) Then
       s = s & " and dbo.project_billl.bill_date <=" & SQLDate(ToDate10, True) & " "
    End If
    
     If Check2.value = vbChecked Then
      s = s & "   and isnull( project_billl.zatcaStatus,0)=1   and warrningmessage<>''   "
    Else
        s = s & "   and isnull( project_billl.zatcaStatus,0)=1   "
    
    End If
    
    
       If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     project_billl.branch_no In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
        
                
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and     project_billl.branch_no = " & val(DcBranches(3).BoundText)
        
        
    End If
           
            
            
            
            
            
    
'    s = s & " SELECT tblActivitesType.id as ActivityTypeId, TblBranchesData.branch_id,TblBranchesData.branch_name BranchName, tblActivitesType.Name as ActivityName, "
'    s = s & "     TblHandWages.Invoicetype"
'   s = s & " ,dbo.TblHandWages.ErrorMessageS,chkTaxExempt = 0"
'   s = s & " ,dbo.TblHandWages.RecordDate DateBaptizing"
'   s = s & " ,CAST(dbo.TblHandWages.NoteSerial1 AS VARCHAR(10)) order_no"
'   s = s & " ,order_no ReturnSerial"
'   s = s & " ,dbo.TblHandWages.RecordDate SalesInvoiceDate"
'   s = s & " ,dbo.TblHandWages.id Transaction_ID"
'   s = s & " ,PayeeFinancialAccount = (SELECT"
'    s = s & "             IBan"
'    s = s & "      From BanksData"
'    s = s & "         WHERE bankid = 0)"
'   s = s & " ,dbo.TblHandWages.Noteid AS Expr1"
'    s = s & " ,10 as Transaction_Type"
'
'   s = s & " ,dbo.TblHandWages.NoteSerial1 AS id"
'   s = s & " ,dbo.TblHandWages.RecordDate AS IssueDate"
'   s = s & " ,dbo.TblHandWages.RecTime AS IssueTim"
'   s = s & " ,dbo.TblHandWages.InvoiceTypeCodeID"
'   s = s & " ,dbo.TblHandWages.InvoiceTypeCodename"
'   s = s & " ,dbo.TblHandWages.DocumentCurrencyCode"
'   s = s & " ,dbo.TblHandWages.TaxCurrencyCode"
'   s = s & " ,dbo.TblHandWages.InvoiceDocumentReferenceID"
'   s = s & " ,dbo.TblHandWages.AdditionalDocumentReferenceICVUUID"
'   s = s & " ,dbo.TblHandWages.ActualDeliveryDate"
'   s = s & " ,dbo.TblHandWages.LatestDeliveryDate"
'   s = s & " ,dbo.TblHandWages.PaymentMeansCode"
'   s = s & " ,dbo.TblHandWages.InstructionNote"
'   s = s & " ,dbo.TblHandWages.paymentnote"
'   s = s & " ,dbo.TblCustemers.CustGID AS Identificationid"
'   s = s & " ,'CRN' AS schemeID"
'   s = s & " ,dbo.TblCustemers.StreetName"
'   s = s & " ,dbo.TblCustemers.AdditionalStreetName"
'   s = s & " ,dbo.TblCustemers.BuildingNumber"
'   s = s & " ,dbo.TblCustemers.PlotIdentification"
'   s = s & " ,dbo.TblCustemers.CityName"
'   s = s & " ,dbo.TblCustemers.PostalZone"
'   s = s & " ,dbo.TblCustemers.CountrySubentity"
'   s = s & " ,dbo.TblCustemers.CitySubdivisionName"
'   s = s & " ,dbo.TblCustemers.IdentificationCode"
'   s = s & " ,dbo.TblCustemers.CusNamee AS RegistrationName"
'   s = s & " ,dbo.TblCustemers.VATNO AS CompanyID , dbo.TblCustemers.Id700 "
'   s = s & " ,0 AS allowancechargeAmount"
'   s = s & " ,'Discount' AS AllowanceChargeReason"
'   s = s & " ,'S' AS TaxCategoryID"
'   s = s & " ,'15' AS TaxCategoryPercent"
'   s = s & " ,TblHandWages.last_changed"
'
'   s = s & " ,Net AS PayableAmount"
'   s = s & " ,0 AS PrepaidAmount"
'   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
'   s = s & " ,dbo.transactionsVatDetails.PIH"
'   s = s & " ,dbo.transactionsVatDetails.QRCode"
'   s = s & " ,dbo.transactionsVatDetails.UUID"
'   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
'   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
'   s = s & " ,dbo.transactionsVatDetails.SingedXML"
'   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
'   s = s & " ,3 AS DocType,vat2 as VatValue,TblCustemers.Export"
''    s = s & " From dbo.TblCustemers"
''s = s & " INNER JOIN dbo.TblHandWages"
''s = s & "     ON dbo.TblCustemers.CusID = dbo.TblHandWages.CusID"
''s = s & " left outer JOIN dbo.transactionsVatDetails"
''s = s & " ON dbo.TblHandWages.ID = dbo.transactionsVatDetails.Transaction_ID and transactionsVatDetails.TableName = 'TblHandWages'"
''s = s & "         AND ISNULL(transactionsVatDetails.isdeleted, 0) = 0"
''s = s & " Inner join TblBranchesData On TblBranchesData.branch_id = TblHandWages.BranchId inner join tblActivitesType On tblActivitesType.id =TblBranchesData.ActivityTypeId  "




 s = s & "  Union all"


  s = s & "  SELECT tblActivitesType.id as ActivityTypeId,TblBranchesData.branch_id,TblBranchesData.branch_name BranchName, tblActivitesType.Name as ActivityName, TblHandWages.Invoicetype"
   s = s & " ,dbo.TblHandWages.warrningmessage,chkTaxExempt =0"
   s = s & " ,dbo.TblHandWages.id Transaction_ID"
   s = s & " ,PayeeFinancialAccount = (SELECT"
    s = s & "             IBan"
    s = s & "         From BanksData"
    s = s & "      WHERE bankid = 0)"
    s = s & "    ,dbo.TblHandWages.id AS Expr1"
   s = s & " ,4 Transaction_Type"
   
   s = s & " ,CAST(dbo.TblHandWages.NoteSerial1 AS NVARCHAR(50)) AS id"
   s = s & " ,dbo.TblHandWages.RecordDate AS IssueDate"
   s = s & " ,dbo.TblHandWages.RecTime AS IssueTim"
   s = s & " ,dbo.TblHandWages.InvoiceTypeCodeID"
   s = s & " ,dbo.TblHandWages.InvoiceTypeCodename"
   s = s & " ,dbo.TblHandWages.DocumentCurrencyCode"
   s = s & " ,dbo.TblHandWages.TaxCurrencyCode"
   s = s & " ,dbo.TblHandWages.InvoiceDocumentReferenceID"
   s = s & " ,dbo.TblHandWages.AdditionalDocumentReferenceICVUUID"
   s = s & " ,dbo.TblHandWages.ActualDeliveryDate"
   s = s & " ,dbo.TblHandWages.LatestDeliveryDate"
   s = s & " ,dbo.TblHandWages.PaymentMeansCode"
   s = s & " ,dbo.TblHandWages.InstructionNote"
   s = s & " ,dbo.TblHandWages.paymentnote"
   s = s & " ,dbo.TblCustemers.CustGID AS Identificationid"
   s = s & " ,'CRN' AS schemeID"
   s = s & " ,dbo.TblCustemers.StreetName"
   s = s & " ,dbo.TblCustemers.AdditionalStreetName"
   s = s & " ,dbo.TblCustemers.BuildingNumber"
   s = s & " ,dbo.TblCustemers.PlotIdentification"
   s = s & " ,dbo.TblCustemers.CityName"
   s = s & " ,dbo.TblCustemers.PostalZone"
   s = s & " ,dbo.TblCustemers.CountrySubentity"
   s = s & " ,dbo.TblCustemers.CitySubdivisionName"
   s = s & " ,dbo.TblCustemers.IdentificationCode"
   s = s & " ,dbo.TblCustemers.CusNamee AS RegistrationName"
   s = s & " ,dbo.TblCustemers.VATNO AS CompanyID, dbo.TblCustemers.Id700 "
   
   's = s & " ,dbo.project_billl.PerforValue +project_billl.DiscountGMater + project_billl.Discount4 AS allowancechargeAmount"
   's = s & " ,project_billl.discount + project_billl.DiscountGMater + project_billl.advancedPayment AS allowancechargeAmount"
   s = s & " ,TblHandWages.TotalDisc  AS allowancechargeAmount"
   s = s & " ,'Discount' AS AllowanceChargeReason"
   s = s & " ,'S' AS TaxCategoryID"
   s = s & " ,'15' AS TaxCategoryPercent"
   s = s & " ,TblHandWages.last_changed"
   s = s & " ,TblHandWages.TotalNet AS PayableAmount "
  ' s = s & " ,project_billl.total AS PayableAmount "
  '  s = s & " ,project_billl.TotalValue  AS PayableAmount"
   s = s & " ,0 AS PrepaidAmount"
   s = s & " ,dbo.transactionsVatDetails.SingedXMLFileName"
   s = s & " ,dbo.transactionsVatDetails.PIH"
   s = s & " ,dbo.transactionsVatDetails.QRCode"
   s = s & " ,dbo.transactionsVatDetails.UUID"
   s = s & " ,dbo.transactionsVatDetails.InvoiceHash"
   s = s & " ,dbo.transactionsVatDetails.EncodedInvoice"
   s = s & " ,dbo.transactionsVatDetails.SingedXML"
   s = s & " ,dbo.transactionsVatDetails.QrCodeDataPath"
   s = s & " ,5 AS DocType"
   
s = s & " From transactionsVatDetails"
s = s & " RIGHT OUTER JOIN TblHandWages"
s = s & "     ON transactionsVatDetails.Transaction_ID = TblHandWages.ID"
s = s & " LEFT OUTER JOIN TblCustemers"
s = s & "     ON transactionsVatDetails.TableName = 'TblHandWages'"
s = s & "         AND ISNULL(transactionsVatDetails.IsDeleted, 0) = 0"
s = s & "         AND TblHandWages.CusID = TblCustemers.CusID"
s = s & " LEFT OUTER JOIN TblBranchesData"
s = s & "     ON TblBranchesData.branch_id = TblHandWages.BranchID"
s = s & " LEFT OUTER JOIN tblActivitesType"
s = s & "     ON tblActivitesType.id = TblBranchesData.ActivityTypeId"


            
            
 s = s & " Where 1 = 1 "
    s = s & "   and isnull( TblHandWages.zatcaStatus,0)=1   "
    If Not IsNull(FromDate10.value) Then
       s = s & " and dbo.TblHandWages.RecordDate >=" & SQLDate(FromDate10, True) & " "
    End If
    
    If Not IsNull(ToDate10.value) Then
       s = s & " and dbo.TblHandWages.RecordDate <=" & SQLDate(ToDate10, True) & " "
    End If
    
     If Check2.value = vbChecked Then
      s = s & "   and isnull( TblHandWages.zatcaStatus,0)=1   and warrningmessage<>''   "
    Else
        s = s & "   and isnull( TblHandWages.zatcaStatus,0)=1   "
    
    End If
    
    
       If Trim(DcBranches(2).Text) <> "" Then
         s = s & "  and     TblHandWages.BranchID In "
        s = s & " (Select TblBranchesData.branch_id From dbo.TblBranchesData"
        s = s & " Where (ActivityTypeId = " & DcBranches(2).BoundText & "))"
  End If
        
                
    If Trim(DcBranches(3).Text) <> "" Then
         s = s & "  and     TblHandWages.BranchID = " & val(DcBranches(3).BoundText)
        
        
    End If
                       
            
            
            
     s = s & "  Union all"

s = s & "  SELECT tblActivitesType.id AS ActivityTypeId, TblBranchesData.branch_id, TblBranchesData.branch_name AS BranchName, tblActivitesType.Name AS ActivityName, "
s = s & "     tblContractInsAllocationsDetails.Invoicetype, "
s = s & "     tblContractInsAllocationsDetails.warrningmessage, chkTaxExempt = 0, "
s = s & "     tblContractInsAllocationsDetails.id AS Transaction_ID, "
s = s & "     PayeeFinancialAccount = (SELECT IBan FROM BanksData WHERE BankID = 0), "
s = s & "     tblContractInsAllocationsDetails.id AS Expr1, "
s = s & "     5 AS Transaction_Type, "

s = s & " CAST(dbo.tblContractInsAllocationsDetails.NoteSerial1H AS NVARCHAR(50)) AS id,"
s = s & "     tblContractInsAllocationsDetails.DateRec AS IssueDate, "
s = s & "     tblContractInsAllocationsDetails.RecTime AS IssueTim, "
s = s & "     tblContractInsAllocationsDetails.InvoiceTypeCodeID, "
s = s & "     tblContractInsAllocationsDetails.InvoiceTypeCodename, "
s = s & "     tblContractInsAllocationsDetails.DocumentCurrencyCode, "
s = s & "     tblContractInsAllocationsDetails.TaxCurrencyCode, "
s = s & "     tblContractInsAllocationsDetails.InvoiceDocumentReferenceID, "
s = s & "     tblContractInsAllocationsDetails.AdditionalDocumentReferenceICVUUID, "
s = s & "     tblContractInsAllocationsDetails.ActualDeliveryDate, "
s = s & "     tblContractInsAllocationsDetails.LatestDeliveryDate, "
s = s & "     tblContractInsAllocationsDetails.PaymentMeansCode, "
s = s & "     tblContractInsAllocationsDetails.InstructionNote, "
s = s & "     tblContractInsAllocationsDetails.paymentnote, "
s = s & "     TblCustemers.CustGID AS Identificationid, "
s = s & "     'CRN' AS schemeID, "
s = s & "     TblCustemers.StreetName, "
s = s & "     TblCustemers.AdditionalStreetName, "
s = s & "     TblCustemers.BuildingNumber, "
s = s & "     TblCustemers.PlotIdentification, "
s = s & "     TblCustemers.CityName, "
s = s & "     TblCustemers.PostalZone, "
s = s & "     TblCustemers.CountrySubentity, "
s = s & "     TblCustemers.CitySubdivisionName, "
s = s & "     TblCustemers.IdentificationCode, "
s = s & "     TblCustemers.CusNamee AS RegistrationName, "
s = s & "     TblCustemers.VATNO AS CompanyID, TblCustemers.Id700, "
s = s & "     0 AS allowancechargeAmount, "
s = s & "     'Discount' AS AllowanceChargeReason, "
s = s & "     'S' AS TaxCategoryID, "
s = s & "     '15' AS TaxCategoryPercent, "
s = s & "     tblContractInsAllocationsDetails.last_changed, "
s = s & "     tblContractInsAllocationsDetails.installValue + isnull(VATValue,0) AS PayableAmount, "
s = s & "     0 AS PrepaidAmount, "
s = s & "     transactionsVatDetails.SingedXMLFileName, "
s = s & "     transactionsVatDetails.PIH, "
s = s & "     transactionsVatDetails.QRCode, "
s = s & "     transactionsVatDetails.UUID, "
s = s & "     transactionsVatDetails.InvoiceHash, "
s = s & "     transactionsVatDetails.EncodedInvoice, "
s = s & "     transactionsVatDetails.SingedXML, "
s = s & "     transactionsVatDetails.QrCodeDataPath, "
s = s & "     6 AS DocType "

s = s & " FROM tblActivitesType "
s = s & " RIGHT OUTER JOIN TblContract "
s = s & " LEFT OUTER JOIN TblBranchesData ON TblContract.Branch_NO = TblBranchesData.branch_id ON tblActivitesType.id = TblBranchesData.ActivityTypeId "
s = s & " RIGHT OUTER JOIN tblContractInsAllocationsDetails "
s = s & " INNER JOIN TblCustemers ON tblContractInsAllocationsDetails.CusID = TblCustemers.CusID ON TblContract.ContNo = tblContractInsAllocationsDetails.ContNo "
s = s & " LEFT OUTER JOIN transactionsVatDetails ON tblContractInsAllocationsDetails.id = transactionsVatDetails.Transaction_ID "
s = s & "     AND transactionsVatDetails.TableName = 'tblContractInsAllocationsDetails' "
s = s & "     AND ISNULL(transactionsVatDetails.IsDeleted, 0) = 0 "

s = s & " WHERE 1 = 1 "

If Not IsNull(FromDate10.value) Then
    s = s & " AND tblContractInsAllocationsDetails.Installdate >= " & SQLDate(FromDate10, True) & " "
End If

If Not IsNull(ToDate10.value) Then
    s = s & " AND tblContractInsAllocationsDetails.Installdate <= " & SQLDate(ToDate10, True) & " "
End If

If Check2.value = vbChecked Then
    s = s & " AND ISNULL(tblContractInsAllocationsDetails.zatcaStatus, 0) = 1 AND warrningmessage <> '' "
Else
    s = s & " AND ISNULL(tblContractInsAllocationsDetails.zatcaStatus, 0) = 1 "
End If

If Trim(DcBranches(2).Text) <> "" Then
    s = s & " AND TblBranchesData.branch_id IN (SELECT branch_id FROM TblBranchesData WHERE ActivityTypeId = " & DcBranches(2).BoundText & ") "
End If

If Trim(DcBranches(3).Text) <> "" Then
    s = s & " AND TblBranchesData.branch_id = " & val(DcBranches(3).BoundText) & " "
End If
       
            
        
    s = s & "                      ORDER by  InvoiceTypeCodeID desc ,Transaction_Date, Id"
    
     Set rsDummy = New ADODB.Recordset
    
If SystemOptions.CanUploadZakatOpt Then
    ConectionFirst
    
    
            s = ""
            s = s & " SELECT "
            s = s & " PropertyDueBatchDetail.Id AS Transaction_ID,ActivityTypeId = 0,Id = NoteSerial1H,SingedXMLFileName ='', Invoicetype = 1,ActivityName = '',"
            s = s & " 'RS Contract' AS TypeName, "
            s = s & " ISNULL(PropertyDueBatchDetail.PaymentMeansCode, '30') AS PaymentMeansCode, "
s = s & " ISNULL(PropertyDueBatchDetail.paymentnote, 'Payment by Credit') AS paymentnote, "
s = s & " ISNULL(PropertyDueBatchDetail.RecTime, CONVERT(nvarchar(20), GETDATE(), 108)) AS RecTime, "
s = s & " ISNULL(PropertyDueBatchDetail.TableName, 'PropertyDueBatchDetail') AS TableName, "
s = s & " ISNULL(PropertyDueBatchDetail.DocumentCurrencyCode, 'SAR') AS DocumentCurrencyCode, "
s = s & " ISNULL(PropertyDueBatchDetail.TaxCurrencyCode, 'SAR') AS TaxCurrencyCode, "
s = s & " ISNULL(PropertyDueBatchDetail.ActualDeliveryDate, CAST(GETDATE() AS date)) AS ActualDeliveryDate, "
s = s & " ISNULL(PropertyDueBatchDetail.LatestDeliveryDate, CAST(GETDATE() AS date)) AS LatestDeliveryDate, "

s = s & " ISNULL(PropertyDueBatchDetail.InvoiceTypeCodeID, 388) AS InvoiceTypeCodeID, "
s = s & " CASE "
s = s & "     WHEN PropertyContractBatch.BatchTotal >= 1000 " ' €Ì— BatchTotal ·«”„ «·⁄„Êœ «·’ÕÌÕ ·Ê „Œ ·›
s = s & "         THEN CASE WHEN PropertyDueBatchDetail.Export = 1 THEN '0100100' ELSE '0100000' END "
s = s & "     ELSE '0200000' "
s = s & " END AS InvoiceTypeCodename, "

            s = s & " PropertyContractBatch.BatchDate AS IssueDate, "
            s = s & " PropertyContractBatch.BatchDate AS SalesInvoiceDate,PayeeFinancialAccount = '', "
            s = s & " PropertyContractBatch.BatchTotal AS PayableAmount, "
            s = s & " PropertyContractBatch.BatchTotal AS PrepaidAmount, "
            s = s & " PropertyRenter.ArName AS RegistrationName, "
            s = s & " PropertyRenter.VATNo AS CompanyID, "
            s = s & " PropertyRenter.RegistrationNo AS Identificationid, "
            s = s & " PropertyRenter.Email, "
            s = s & " PropertyRenter.Address AS StreetName, "
            s = s & " PropertyContractBatch.Id AS BatchID, "
            s = s & " 10 AS Transaction_Type, "
            s = s & " '' AS order_no, "
            s = s & " '' AS ReturnSerial, "
            s = s & " '' AS AllowanceChargeReason, "
            s = s & " 0 AS allowancechargeAmount, "
            s = s & " PropertyDueBatchDetail.ErrorMessageS, "
            s = s & " 0 AS chkTaxExempt, "
            s = s & " PropertyDueBatchDetail.NoteSerial1H, "
            
            
            s = s & " PropertyDueBatchDetail.InvoiceDocumentReferenceID, "
            s = s & " PropertyDueBatchDetail.AdditionalDocumentReferenceICVUUID, "
           
            
            s = s & " PropertyDueBatchDetail.InstructionNote, "
            
            s = s & " 'CRN' AS schemeID, "
            s = s & " PropertyRenter.Address AS AdditionalStreetName, "
            s = s & " NULL AS BuildingNumber, "
            s = s & " NULL AS PlotIdentification, "
            s = s & " 'Riyadh' AS CityName, "
            s = s & " '12345' AS PostalZone, "
            s = s & " NULL AS CountrySubentity, "
            s = s & " NULL AS CitySubdivisionName, "
            s = s & " NULL AS IdentificationCode, "
            s = s & " NULL AS Id700, "
            
           ' ≈Ã„«·Ì «·÷—Ì»… «·„œ›Ê⁄… ⁄»— ﬂ· «·»‰Êœ
            s = s & " 0 AS FATValue, "
            s = s & " CAST(ROUND( "
            s = s & "       ISNULL(PropertyContractBatch.BatchRentValueTaxes,0) "
            s = s & "     + ISNULL(PropertyContractBatch.BatchWaterValueTaxes,0) "
            s = s & "     + ISNULL(PropertyContractBatch.BatchElectricityValueTaxes,0) "
            s = s & "     + ISNULL(PropertyContractBatch.BatchCommissionValueTaxes,0) "
            s = s & "     + ISNULL(PropertyContractBatch.BatchGasValueTaxes,0) "
            s = s & "     + ISNULL(PropertyContractBatch.BatchServicesValueTaxes,0) "
            s = s & "     + ISNULL(PropertyContractBatch.BatchInsuranceValueTaxes,0) "
            s = s & " , 2) AS DECIMAL(18,2)) AS VatValue, "
 
            s = s & " STUFF( "
            s = s & "     (CASE WHEN ISNULL(PropertyContractBatch.BatchRentValueTaxes,0) > 0 "
            s = s & "           THEN N' | ÷—Ì»… «·≈ÌÃ«— ' + CONVERT(NVARCHAR(30), CAST(ROUND(ISNULL(PropertyContractBatch.BatchRentValueTaxes,0),2) AS DECIMAL(18,2))) "
            s = s & "           ELSE N'' END) "
            s = s & "   + (CASE WHEN ISNULL(PropertyContractBatch.BatchWaterValueTaxes,0) > 0 "
            s = s & "           THEN N' | ÷—Ì»… «·„Ì«Â ' + CONVERT(NVARCHAR(30), CAST(ROUND(ISNULL(PropertyContractBatch.BatchWaterValueTaxes,0),2) AS DECIMAL(18,2))) "
            s = s & "           ELSE N'' END) "
            s = s & "   + (CASE WHEN ISNULL(PropertyContractBatch.BatchElectricityValueTaxes,0) > 0 "
            s = s & "           THEN N' | ÷—Ì»… «·ﬂÂ—»«¡ ' + CONVERT(NVARCHAR(30), CAST(ROUND(ISNULL(PropertyContractBatch.BatchElectricityValueTaxes,0),2) AS DECIMAL(18,2))) "
            s = s & "           ELSE N'' END) "
            s = s & "   + (CASE WHEN ISNULL(PropertyContractBatch.BatchCommissionValueTaxes,0) > 0 "
            s = s & "           THEN N' | ÷—Ì»… «·”⁄Ì ' + CONVERT(NVARCHAR(30), CAST(ROUND(ISNULL(PropertyContractBatch.BatchCommissionValueTaxes,0),2) AS DECIMAL(18,2))) "
            s = s & "           ELSE N'' END) "
            s = s & "   + (CASE WHEN ISNULL(PropertyContractBatch.BatchGasValueTaxes,0) > 0 "
            s = s & "           THEN N' | ÷—Ì»… «·€«“ ' + CONVERT(NVARCHAR(30), CAST(ROUND(ISNULL(PropertyContractBatch.BatchGasValueTaxes,0),2) AS DECIMAL(18,2))) "
            s = s & "           ELSE N'' END) "
            s = s & "   + (CASE WHEN ISNULL(PropertyContractBatch.BatchServicesValueTaxes,0) > 0 "
            s = s & "           THEN N' | ÷—Ì»… «·Œœ„«  ' + CONVERT(NVARCHAR(30), CAST(ROUND(ISNULL(PropertyContractBatch.BatchServicesValueTaxes,0),2) AS DECIMAL(18,2))) "
            s = s & "           ELSE N'' END) "
            s = s & "   + (CASE WHEN ISNULL(PropertyContractBatch.BatchInsuranceValueTaxes,0) > 0 "
            s = s & "           THEN N' | ÷—Ì»… «· √„Ì‰ ' + CONVERT(NVARCHAR(30), CAST(ROUND(ISNULL(PropertyContractBatch.BatchInsuranceValueTaxes,0),2) AS DECIMAL(18,2))) "
            s = s & "           ELSE N'' END), "
            s = s & "   1, 3, N'' "
            s = s & " ) AS VatBreakdownNote, "

            
            s = s & " NULL AS last_changed, "
            s = s & " NULL AS IssueTim, "
            s = s & " NULL AS PIH, "
            s = s & " NULL AS QRCode, "
            s = s & " NULL AS UUID, "
            s = s & " NULL AS InvoiceHash, "
            s = s & " NULL AS EncodedInvoice, "
            s = s & " NULL AS SingedXML, "
            
            
            
            s = s & "     CASE WHEN IsNull(PropertyContract.VATPercentage,0) = 0 THEN 'O' ELSE 'S' END AS TaxCategoryID, "
s = s & "     IsNull(PropertyContract.VATPercentage,0) * 100 AS TaxCategoryPercent, "


s = s & "     PropertyDueBatchDetail.last_changed, "
s = s & "     PropertyContractBatch.BatchTotal AS PayableAmount, "
s = s & "     0 AS PrepaidAmount, "
s = s & "     transactionsVatDetails.SingedXMLFileName, "
s = s & "     transactionsVatDetails.PIH, "
s = s & "     transactionsVatDetails.QRCode, "
s = s & "     transactionsVatDetails.UUID, "
s = s & "     transactionsVatDetails.InvoiceHash, "
s = s & "     transactionsVatDetails.EncodedInvoice, "
s = s & "     transactionsVatDetails.SingedXML, "
s = s & "     transactionsVatDetails.QrCodeDataPath, "

            
            s = s & " 6 AS DocType, "
            s = s & " 0 AS Export, "
            s = s & " PropertyDueBatchDetail.QrCodeData, "
            s = s & " PropertyDueBatchDetail.QrCodeImage, "
            s = s & " PropertyDueBatchDetail.zatcaStatus, "
            s = s & " PropertyDueBatchDetail.warrningmessage, "
            s = s & " PropertyDueBatchDetail.Iban, "
            
            s = s & "     Department.id as branch_id,"
            s = s & " Department.ArName as  branch_name,"
            s = s & " Department.ArName as  branchname"
            s = s & " FROM PropertyDueBatchDetail "
            s = s & " INNER JOIN PropertyContractBatch ON PropertyContractBatch.Id = PropertyDueBatchDetail.PropertyContractBatchId "
            s = s & " INNER JOIN PropertyContract ON PropertyContract.Id = PropertyContractBatch.MainDocId "
            s = s & " INNER JOIN PropertyRenter ON PropertyContract.PropertyRenterId = PropertyRenter.Id "
            s = s & " Left OUTER join Department On PropertyContract.DepartmentId = Department.Id"
            s = s & "    Left OUTER join transactionsVatDetails ON transactionsVatDetails.Transaction_ID = PropertyDueBatchDetail.ID"
            s = s & "     and transactionsVatDetails.TableName = 'PropertyDueBatchDetail'"
            s = s & "         AND ISNULL(transactionsVatDetails.IsDeleted, 0) = 0"
            s = s & " WHERE 1 = 1 and IsNull(IsSelected,0) = 1"
          '  s = s & "     and PropertyContract.id not in (Select PropertyContractTermination.PropertyContractId from PropertyContractTermination)"
            ' ›· —… «· «—ÌŒ
            If Not IsNull(FromDate.value) Then
                s = s & " AND PropertyContractBatch.BatchDate >= " & SQLDate(FromDate10, True) & " "
            End If
            
            If Not IsNull(ToDate.value) Then
                s = s & " AND PropertyContractBatch.BatchDate <= " & SQLDate(ToDate10, True) & " "
            End If
            
            ' ›· — Õ«·… ZATCA
            s = s & " AND ISNULL(PropertyDueBatchDetail.zatcaStatus, 0) = 1 "
            
            ' ›·« — «·›—Ê⁄ (·Ê ⁄‰œﬂ ÃœÊ· Œ«’ ··›—Ê⁄ √Ê —»ÿ „‰«”»)
'            If Trim(DcBranches(2).text) <> "" Then
'                s = s & " AND PropertyDueBatchDetail.branch_id IN (SELECT branch_id FROM PropertyDueBatchDetail WHERE ActivityTypeId = " & DcBranches(2).BoundText & ") "
'            End If
'
'            If Trim(DcBranches(3).text) <> "" Then
'                s = s & " AND PropertyDueBatchDetail.branch_id = " & val(DcBranches(3).BoundText) & " "
'            End If
            
            s = s & " ORDER BY PropertyDueBatchDetail.InvoiceTypeCodeID DESC, PropertyContractBatch.BatchDate, PropertyDueBatchDetail.NoteSerial1H "
            
            
            
          '
          
          
          s = ""

'/*==================== CTE: ⁄ﬁÊœ RS ====================*/
SqlAdd s, "WITH Base_RS AS ("
SqlAdd s, "  SELECT"
SqlAdd s, "    CAST(0 AS int)                          AS ActivityTypeId,"
SqlAdd s, "    N'RS Contract'                           AS TypeName,"
SqlAdd s, "    dep.Id                                   AS branch_id,"
SqlAdd s, "    dep.ArName                               AS BranchName,"
SqlAdd s, "    N''                                      AS ActivityName,"
SqlAdd s, "    CAST(388 AS int)                         AS Invoicetype,"
SqlAdd s, "    ISNULL(pdbd.ErrorMessageS,N'')           AS ErrorMessageS,"
SqlAdd s, "    CAST(0 AS bit)                           AS chkTaxExempt,"
SqlAdd s, "    pcb.BatchDate                            AS DateBaptizing,"
SqlAdd s, "    N''                                      AS order_no,"
SqlAdd s, "    N''                                      AS ReturnSerial,"
SqlAdd s, "    pcb.BatchDate                            AS SalesInvoiceDate,"
SqlAdd s, "    CAST(pdbd.Id AS bigint)                  AS Transaction_ID,"
SqlAdd s, "    N''                                      AS PayeeFinancialAccount,"
SqlAdd s, "    CAST(pdbd.Id AS bigint)                  AS Expr1,"
SqlAdd s, "    CAST(10 AS int)                          AS Transaction_Type,"
SqlAdd s, "    CAST(pdbd.NoteSerial1H AS nvarchar(50))  AS id,"
SqlAdd s, "    pcb.BatchDate                            AS IssueDate,"
SqlAdd s, "    ISNULL(CONVERT(nvarchar(8),GETDATE(),108),N'') AS IssueTim,"
SqlAdd s, "    ISNULL(pdbd.InvoiceTypeCodeID,388)       AS InvoiceTypeCodeID,"
SqlAdd s, "    CASE WHEN pcb.BatchTotal>=1000"
SqlAdd s, "         THEN CASE WHEN ISNULL(pdbd.Export,0)=1 THEN '0100100' ELSE '0100000' END"
SqlAdd s, "         ELSE '0200000' END                  AS InvoiceTypeCodename,"
SqlAdd s, "    ISNULL(pdbd.DocumentCurrencyCode,N'SAR') AS DocumentCurrencyCode,"
SqlAdd s, "    ISNULL(pdbd.TaxCurrencyCode,N'SAR')      AS TaxCurrencyCode,"
SqlAdd s, "    ISNULL(pdbd.InvoiceDocumentReferenceID,N'')     AS InvoiceDocumentReferenceID,"
SqlAdd s, "    ISNULL(pdbd.AdditionalDocumentReferenceICVUUID,N'') AS AdditionalDocumentReferenceICVUUID,"
SqlAdd s, "    ISNULL(pdbd.ActualDeliveryDate,CAST(GETDATE() AS date))  AS ActualDeliveryDate,"
SqlAdd s, "    ISNULL(pdbd.LatestDeliveryDate,CAST(GETDATE() AS date))  AS LatestDeliveryDate,"
SqlAdd s, "    ISNULL(CAST(pdbd.PaymentMeansCode AS float),30)          AS PaymentMeansCode,"
SqlAdd s, "    ISNULL(pdbd.InstructionNote,N'')                 AS InstructionNote,"
SqlAdd s, "    ISNULL(pdbd.paymentnote,N'Payment by Credit')    AS paymentnote,"
SqlAdd s, "    pr.RegistrationNo                       AS Identificationid,"
SqlAdd s, "    N'CRN'                                   AS schemeID,"
SqlAdd s, "    pr.Address                               AS StreetName,"
SqlAdd s, "    pr.Address                               AS AdditionalStreetName,"
SqlAdd s, "    NULL                                     AS BuildingNumber,"
SqlAdd s, "    NULL                                     AS PlotIdentification,"
SqlAdd s, "    N'Riyadh'                                AS CityName,"
SqlAdd s, "    N'12345'                                 AS PostalZone,"
SqlAdd s, "    NULL                                     AS CountrySubentity,"
SqlAdd s, "    NULL                                     AS CitySubdivisionName,"
SqlAdd s, "    NULL                                     AS IdentificationCode,"
SqlAdd s, "    pr.ArName                                AS RegistrationName,"
SqlAdd s, "    pr.VATNo                                 AS CompanyID,"
SqlAdd s, "    NULL                                     AS Id700,"
SqlAdd s, "    CAST(0 AS float)                         AS allowancechargeAmount,"
SqlAdd s, "    N'Discount'                              AS AllowanceChargeReason,"
SqlAdd s, "    CASE WHEN ISNULL(pc.VATPercentage,0)=0 THEN 'O' ELSE 'S' END AS TaxCategoryID,"
SqlAdd s, "    ISNULL(pc.VATPercentage,0)*100           AS TaxCategoryPercent,"
SqlAdd s, "    pdbd.last_changed                        AS last_changed,"
SqlAdd s, "    pcb.BatchTotal                           AS PayableAmount,"
SqlAdd s, "    CAST(0 AS float)                         AS PrepaidAmount,"
SqlAdd s, "    ISNULL(tvd.SingedXMLFileName,N'')        AS SingedXMLFileName,"
SqlAdd s, "    ISNULL(tvd.PIH,N'')                      AS PIH,"
SqlAdd s, "    ISNULL(tvd.QRCode,N'')                   AS QRCode,"
SqlAdd s, "    ISNULL(tvd.UUID,N'')                     AS UUID,"
SqlAdd s, "    ISNULL(tvd.InvoiceHash,N'')              AS InvoiceHash,"
SqlAdd s, "    ISNULL(tvd.EncodedInvoice,N'')           AS EncodedInvoice,"
SqlAdd s, "    ISNULL(tvd.SingedXML,N'')                AS SingedXML,"
SqlAdd s, "    ISNULL(tvd.QrCodeDataPath,N'')           AS QrCodeDataPath,"
SqlAdd s, "    CAST(6 AS int)                           AS DocType,"
SqlAdd s, "    CAST(ROUND("
SqlAdd s, "        ISNULL(pcb.BatchRentValueTaxes,0)+ISNULL(pcb.BatchWaterValueTaxes,0)+"
SqlAdd s, "        ISNULL(pcb.BatchElectricityValueTaxes,0)+ISNULL(pcb.BatchCommissionValueTaxes,0)+"
SqlAdd s, "        ISNULL(pcb.BatchGasValueTaxes,0)+ISNULL(pcb.BatchServicesValueTaxes,0)+"
SqlAdd s, "        ISNULL(pcb.BatchInsuranceValueTaxes,0),2) AS decimal(18,2)) AS VatValue,"
SqlAdd s, "    ISNULL(pdbd.Export,0)                    AS Export,"
SqlAdd s, "    STUFF(("
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchRentValueTaxes,0)>0 THEN N' | ÷—Ì»… «·≈ÌÃ«— ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchRentValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchWaterValueTaxes,0)>0 THEN N' | ÷—Ì»… «·„Ì«Â ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchWaterValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchElectricityValueTaxes,0)>0 THEN N' | ÷—Ì»… «·ﬂÂ—»«¡ ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchElectricityValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchCommissionValueTaxes,0)>0 THEN N' | ÷—Ì»… «·”⁄Ì ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchCommissionValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchGasValueTaxes,0)>0 THEN N' | ÷—Ì»… «·€«“ ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchGasValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchServicesValueTaxes,0)>0 THEN N' | ÷—Ì»… «·Œœ„«  ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchServicesValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END) +"
SqlAdd s, "      (CASE WHEN ISNULL(pcb.BatchInsuranceValueTaxes,0)>0 THEN N' | ÷—Ì»… «· √„Ì‰ ' + CONVERT(nvarchar(30),CAST(ROUND(ISNULL(pcb.BatchInsuranceValueTaxes,0),2) AS decimal(18,2))) ELSE N'' END)"
SqlAdd s, "    ),1,3,N'') AS VatBreakdownNote,warrningmessage"
SqlAdd s, "  FROM PropertyDueBatchDetail pdbd"
SqlAdd s, "  JOIN PropertyContractBatch pcb ON pcb.Id = pdbd.PropertyContractBatchId"
SqlAdd s, "  JOIN PropertyContract pc       ON pc.Id  = pcb.MainDocId"
SqlAdd s, "  JOIN PropertyRenter pr         ON pc.PropertyRenterId = pr.Id"
SqlAdd s, "  LEFT JOIN Department dep       ON pc.DepartmentId = dep.Id"
SqlAdd s, "  LEFT JOIN transactionsVatDetails tvd"
SqlAdd s, "         ON tvd.Transaction_ID = pdbd.Id"
SqlAdd s, "        AND tvd.TableName      = 'PropertyDueBatchDetail'"
SqlAdd s, "        AND ISNULL(tvd.IsDeleted,0)=0"
SqlAdd s, "  WHERE ISNULL(pdbd.IsSelected,0)=1"
SqlAdd s, "    AND ISNULL(pdbd.zatcaStatus,0)=1"   ' ›· — ZATCA = 1 ·‹ RS

' ›·« — «· «—ÌŒ ·‹ RS
If Not IsNull(FromDate.value) Then SqlAdd s, "    AND pcb.BatchDate >= " & SQLDate(FromDate10, True)
If Not IsNull(ToDate.value) Then SqlAdd s, "    AND pcb.BatchDate <= " & SQLDate(ToDate10, True)

SqlAdd s, ")"

'/*==================== CTE: ≈‘⁄«—«  «·Œ’„/«·≈÷«›… ====================*/
SqlAdd s, ", Base_NOTE AS ("
SqlAdd s, "  SELECT"
SqlAdd s, "    ISNULL(dep.ActivityId,0)                  AS ActivityTypeId,"
SqlAdd s, "    N'Credit Or Debit Note (WEB)'             AS TypeName,"
SqlAdd s, "    dep.Id                                    AS branch_id,"
SqlAdd s, "    dep.ArName                                AS BranchName,"
SqlAdd s, "    N''                                       AS ActivityName,"
SqlAdd s, "    CASE d.DebitAndCreditNotificationTypeId WHEN 4 THEN 383 WHEN 5 THEN 381 ELSE ISNULL(CAST(d.InvoiceTypeCodeID AS int),0) END AS Invoicetype,"
SqlAdd s, "    ISNULL(d.ErrorMessageS,N'')               AS ErrorMessageS,"
SqlAdd s, "    CAST(0 AS bit)                            AS chkTaxExempt,"
SqlAdd s, "    d.[Date]                                  AS DateBaptizing,"
SqlAdd s, "    CAST(ISNULL(d.DocumentNumber,N'') AS nvarchar(50)) AS order_no,"
SqlAdd s, "    N''                                       AS ReturnSerial,"
SqlAdd s, "    d.[Date]                                  AS SalesInvoiceDate,"
SqlAdd s, "    CAST(d.Id AS bigint)                      AS Transaction_ID,"
SqlAdd s, "    N''                                       AS PayeeFinancialAccount,"
SqlAdd s, "    CAST(d.Id AS bigint)                      AS Expr1,"
SqlAdd s, "    CASE d.DebitAndCreditNotificationTypeId WHEN 4 THEN 383 WHEN 5 THEN 381 ELSE 0 END AS Transaction_Type,"
SqlAdd s, "    ISNULL(d.InvoiceNo,ISNULL(d.DocumentNumber,N'')) AS id,"
SqlAdd s, "    d.[Date]                                  AS IssueDate,"
SqlAdd s, "    ISNULL(d.RecTime,CONVERT(nvarchar(8),GETDATE(),108)) AS IssueTim,"
SqlAdd s, "    ISNULL(CAST(d.InvoiceTypeCodeID AS int),0) AS InvoiceTypeCodeID,"
SqlAdd s, "    ISNULL(d.InvoiceTypeCodename, CASE WHEN ISNULL(d.MoneyAmount,0)>=1000 THEN '0100000' ELSE '0200000' END) AS InvoiceTypeCodename,"
SqlAdd s, "    ISNULL(d.DocumentCurrencyCode,N'SAR')     AS DocumentCurrencyCode,"
SqlAdd s, "    ISNULL(d.TaxCurrencyCode,N'SAR')          AS TaxCurrencyCode,"
SqlAdd s, "    ISNULL(d.InvoiceDocumentReferenceID,N'')  AS InvoiceDocumentReferenceID,"
SqlAdd s, "    ISNULL(d.AdditionalDocumentReferenceICVUUID,N'') AS AdditionalDocumentReferenceICVUUID,"
SqlAdd s, "    ISNULL(d.ActualDeliveryDate,CAST(GETDATE() AS date)) AS ActualDeliveryDate,"
SqlAdd s, "    ISNULL(d.LatestDeliveryDate,CAST(GETDATE() AS date)) AS LatestDeliveryDate,"
SqlAdd s, "    ISNULL(CAST(d.PaymentMeansCode AS float),30)        AS PaymentMeansCode,"
SqlAdd s, "    ISNULL(d.InstructionNote,N'')            AS InstructionNote,"
SqlAdd s, "    ISNULL(d.paymentnote,N'Payment by Credit') AS paymentnote,"
SqlAdd s, "    COALESCE(r.RegistrationNo,c.RegistrationNo,v.RegistrationNo,N'') AS Identificationid,"
SqlAdd s, "    N'CRN'                                     AS schemeID,"
SqlAdd s, "    N''                                        AS StreetName,"
SqlAdd s, "    N''                                        AS AdditionalStreetName,"
SqlAdd s, "    NULL                                       AS BuildingNumber,"
SqlAdd s, "    NULL                                       AS PlotIdentification,"
SqlAdd s, "    N''                                        AS CityName,"
SqlAdd s, "    N''                                        AS PostalZone,"
SqlAdd s, "    NULL                                       AS CountrySubentity,"
SqlAdd s, "    NULL                                       AS CitySubdivisionName,"
SqlAdd s, "    NULL                                       AS IdentificationCode,"
SqlAdd s, "    COALESCE(r.ArName,c.ArName,v.ArName,d.IssuedPerson,N'') AS RegistrationName,"
SqlAdd s, "    COALESCE(r.VATNo,c.VATNo,v.VATNo,N'')      AS CompanyID,"
SqlAdd s, "    N''                                        AS Id700,"
SqlAdd s, "    CAST(0 AS float)                           AS allowancechargeAmount,"
SqlAdd s, "    N'Discount'                                AS AllowanceChargeReason,"
SqlAdd s, "    CASE WHEN ISNULL(d.VATValue,0)>0 THEN 'S' ELSE 'O' END AS TaxCategoryID,"
SqlAdd s, "    CASE WHEN ISNULL(d.VATPercentage,0)>0 THEN d.VATPercentage*100 ELSE 0 END AS TaxCategoryPercent,"
SqlAdd s, "    d.last_changed                             AS last_changed,"
SqlAdd s, "    ISNULL(d.MoneyAmount,0)+ISNULL(d.VATValue,0) AS PayableAmount,"
SqlAdd s, "    CAST(0 AS float)                           AS PrepaidAmount,"
SqlAdd s, "    N''                                        AS SingedXMLFileName,"
SqlAdd s, "    N''                                        AS PIH,"
SqlAdd s, "    N''                                        AS QRCode,"
SqlAdd s, "    N''                                        AS UUID,"
SqlAdd s, "    N''                                        AS InvoiceHash,"
SqlAdd s, "    N''                                        AS EncodedInvoice,"
SqlAdd s, "    N''                                        AS SingedXML,"
SqlAdd s, "    ISNULL(d.QrCodeDataPath,N'')               AS QrCodeDataPath,"
SqlAdd s, "    CAST(3 AS int)                             AS DocType,"
SqlAdd s, "    ISNULL(CAST(d.VATValue AS float),0)        AS VatValue,"
SqlAdd s, "    ISNULL(d.Export,0)                         AS Export,"
SqlAdd s, "    CAST(N'' AS nvarchar(1000))                AS VatBreakdownNote,warrningmessage"
SqlAdd s, "  FROM dbo.DebitAndCreditNotification d"
SqlAdd s, "  LEFT JOIN dbo.Department dep ON dep.Id = d.DepartmentId"
SqlAdd s, "  LEFT JOIN dbo.Customer   c   ON c.Id  = d.CustomerId"
SqlAdd s, "  LEFT JOIN dbo.Vendor     v   ON v.Id  = d.VendorId"
SqlAdd s, "  LEFT JOIN dbo.PropertyRenter r ON r.Id = d.RenterId"
SqlAdd s, "  WHERE ISNULL(d.zatcaStatus,0)=1"   ' ›· — ZATCA = 1 ··≈‘⁄«—«  √Ì÷«

' ›·« — «· «—ÌŒ ··≈‘⁄«—« 
If Not IsNull(FromDate.value) Then SqlAdd s, "    AND d.[Date] >= " & SQLDate(FromDate10, True)
If Not IsNull(ToDate.value) Then SqlAdd s, "    AND d.[Date] <= " & SQLDate(ToDate10, True)

SqlAdd s, ")"

'/*==================== «·‰ ÌÃ… «·‰Â«∆Ì… ====================*/
SqlAdd s, "SELECT * FROM Base_RS"
SqlAdd s, "UNION ALL"
SqlAdd s, "SELECT * FROM Base_NOTE"
SqlAdd s, "ORDER BY InvoiceTypeCodeID DESC, IssueDate, id;"

            rsDummy.Open s, POSConnection, adOpenKeyset, adLockReadOnly
            Else
            
    
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    
End If
    
   
    grd(2).rows = 1
    grd(2).rows = grd(2).rows + 1
    With grd(2)
      .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
       ' .AutoResize = True
        .AutoSizeMode = flexAutoSizeColWidth
        'Set .WallPaper = Grdback.MoneyWallpaper
  .ColComboList(.ColIndex("viewFile")) = "..."
  .ColComboList(.ColIndex("ErrorMessage")) = "..."
  .ColComboList(.ColIndex("viewQRCode")) = "..."
  
       ' .AutoSize 0,  .Cols - 1, False
 
     
     End With
     
     
         With grd(2)
      .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
       ' .AutoResize = True
        .AutoSizeMode = flexAutoSizeColWidth
        'Set .WallPaper = Grdback.MoneyWallpaper
  .ColComboList(.ColIndex("viewFile")) = "..."
  .ColComboList(.ColIndex("ErrorMessage")) = "..."
  .ColComboList(.ColIndex("viewQRCode")) = "..."
  
       ' .AutoSize 0,  .Cols - 1, False
 
     
     End With
     
     
    Dim i As Long
    Dim mTotalNet As Double
    Dim mTotalDiscountNet As Double
    Dim mTransaction_NetValue As Double
    i = grd(2).rows - 1

 
Dim OtherInformation As New ClsGLOther


    Do While Not rsDummy.EOF
  
         
        Dim e As New ClsGLOther
        e.Invoicetype = IIf(IsNull(rsDummy("Invoicetype").value), 0, rsDummy("Invoicetype").value)
        e.chkTaxExempt = IIf(IsNull(rsDummy("chkTaxExempt").value), 0, rsDummy("chkTaxExempt").value)
  e.ID = IIf(IsNull(rsDummy("id").value), "", rsDummy("id").value)
  e.Id700 = IIf(IsNull(rsDummy("Id700").value), "", rsDummy("Id700").value)
  e.IssueDate = IIf(IsNull(rsDummy("IssueDate").value), "", rsDummy("IssueDate").value)
  e.IssueTim = IIf(IsNull(rsDummy("IssueTim").value), "", rsDummy("IssueTim").value)
 e.InvoiceTypeCodeID = IIf(IsNull(rsDummy("InvoiceTypeCodeID").value), "", rsDummy("InvoiceTypeCodeID").value)
e.InvoiceTypeCodename = IIf(IsNull(rsDummy("InvoiceTypeCodename").value), "", rsDummy("InvoiceTypeCodename").value)
e.DocumentCurrencyCode = IIf(IsNull(rsDummy("DocumentCurrencyCode").value), "", rsDummy("DocumentCurrencyCode").value)
e.TaxCurrencyCode = IIf(IsNull(rsDummy("TaxCurrencyCode").value), "", rsDummy("TaxCurrencyCode").value)
e.InvoiceDocumentReferenceID = IIf(IsNull(rsDummy("InvoiceDocumentReferenceID").value), "", rsDummy("InvoiceDocumentReferenceID").value)
e.AdditionalDocumentReferenceICVUUID = IIf(IsNull(rsDummy("AdditionalDocumentReferenceICVUUID").value), "", rsDummy("AdditionalDocumentReferenceICVUUID").value)
e.ActualDeliveryDate = IIf(IsNull(rsDummy("ActualDeliveryDate").value), "", rsDummy("ActualDeliveryDate").value)
e.LatestDeliveryDate = IIf(IsNull(rsDummy("LatestDeliveryDate").value), "", rsDummy("LatestDeliveryDate").value)
e.PaymentMeansCode = IIf(IsNull(rsDummy("PaymentMeansCode").value), "", rsDummy("PaymentMeansCode").value)
e.InstructionNote = IIf(IsNull(rsDummy("InstructionNote").value), "", rsDummy("InstructionNote").value)
e.PayeeFinancialAccount = IIf(IsNull(rsDummy("PayeeFinancialAccount").value), "", rsDummy("PayeeFinancialAccount").value)
e.paymentnote = IIf(IsNull(rsDummy("paymentnote").value), "", rsDummy("paymentnote").value)
e.Identificationid = IIf(IsNull(rsDummy("Identificationid").value), "", rsDummy("Identificationid").value)
e.schemeID = IIf(IsNull(rsDummy("schemeID").value), "", rsDummy("schemeID").value)
e.StreetName = IIf(IsNull(rsDummy("StreetName").value), "", rsDummy("StreetName").value)
e.AdditionalStreetName = IIf(IsNull(rsDummy("AdditionalStreetName").value), "", rsDummy("AdditionalStreetName").value)
e.BuildingNumber = IIf(IsNull(rsDummy("BuildingNumber").value), "", rsDummy("BuildingNumber").value)
e.PlotIdentification = IIf(IsNull(rsDummy("PlotIdentification").value), "", rsDummy("PlotIdentification").value)
e.CityName = IIf(IsNull(rsDummy("CityName").value), "", rsDummy("CityName").value)
e.PostalZone = IIf(IsNull(rsDummy("PostalZone").value), "", rsDummy("PostalZone").value)
e.CountrySubentity = IIf(IsNull(rsDummy("CountrySubentity").value), "", rsDummy("CountrySubentity").value)
e.CitySubdivisionName = IIf(IsNull(rsDummy("CitySubdivisionName").value), "", rsDummy("CitySubdivisionName").value)
e.IdentificationCode = IIf(IsNull(rsDummy("IdentificationCode").value), "", rsDummy("IdentificationCode").value)
e.RegistrationName = IIf(IsNull(rsDummy("RegistrationName").value), "", rsDummy("RegistrationName").value)
e.CompanyID = IIf(IsNull(rsDummy("CompanyID").value), "", rsDummy("CompanyID").value)
e.allowancechargeAmount = IIf(IsNull(rsDummy("allowancechargeAmount").value), 0, rsDummy("allowancechargeAmount").value)
e.AllowanceChargeReason = IIf(IsNull(rsDummy("AllowanceChargeReason").value), "", rsDummy("AllowanceChargeReason").value)
e.TaxCategoryID = IIf(IsNull(rsDummy("TaxCategoryID").value), "", rsDummy("TaxCategoryID").value)
e.TaxCategoryPercent = IIf(IsNull(rsDummy("TaxCategoryPercent").value), "", rsDummy("TaxCategoryPercent").value)
e.PayableAmount = IIf(IsNull(rsDummy("PayableAmount").value), "", rsDummy("PayableAmount").value)
e.PrepaidAmount = IIf(IsNull(rsDummy("PrepaidAmount").value), 0, rsDummy("PrepaidAmount").value)
  e.Transaction_ID = IIf(IsNull(rsDummy("Transaction_ID").value), "", rsDummy("Transaction_ID").value)
   
   
   
   e.InvoiceHash = IIf(IsNull(rsDummy("InvoiceHash").value), "", rsDummy("InvoiceHash").value)
   
    
                    
   e.SingedXML = IIf(IsNull(rsDummy("SingedXML").value), "", rsDummy("SingedXML").value)
   e.EncodedInvoice = IIf(IsNull(rsDummy("EncodedInvoice").value), "", rsDummy("EncodedInvoice").value)
   e.UUID = IIf(IsNull(rsDummy("UUID").value), "", rsDummy("UUID").value)
   e.QRCode = IIf(IsNull(rsDummy("QRCode").value), "", rsDummy("QRCode").value)
   e.PIH = IIf(IsNull(rsDummy("PIH").value), "", rsDummy("PIH").value)
   e.SingedXMLFileName = IIf(IsNull(rsDummy("SingedXMLFileName").value), "", rsDummy("SingedXMLFileName").value)
   e.docType2 = IIf(IsNull(rsDummy("docType").value), 0, rsDummy("docType").value)
e.QrCodeDataPath = IIf(IsNull(rsDummy("QrCodeDataPath").value), "", rsDummy("QrCodeDataPath").value)
  e.warrningmessage = IIf(IsNull(rsDummy("warrningmessage").value), "", rsDummy("warrningmessage").value)
    grd(2).TextMatrix(i, grd(2).ColIndex("warrningmessage")) = e.warrningmessage
    
        If e.warrningmessage = "" Then
   
    
      
   
   Else
 
        grd(2).cell(flexcpBackColor, i, 1, i, 56) = &H8080FF
              
   
   End If
   
   
   
   
   e.branch_idInvoice = IIf(IsNull(rsDummy("branch_id").value), 0, rsDummy("branch_id").value)
e.ActivityTypeIdInvoice = IIf(IsNull(rsDummy("ActivityTypeId").value), 0, rsDummy("ActivityTypeId").value)

grd(2).TextMatrix(i, grd(2).ColIndex("branch_id")) = rsDummy("branch_id") & ""
grd(2).TextMatrix(i, grd(2).ColIndex("ActivityTypeId")) = rsDummy("ActivityTypeId") & ""
grd(2).TextMatrix(i, grd(2).ColIndex("BranchName")) = rsDummy("BranchName") & ""
grd(2).TextMatrix(i, grd(2).ColIndex("ActivityName")) = rsDummy("ActivityName") & ""
If SystemOptions.CanUploadZakatOpt Then
    grd(2).TextMatrix(i, grd(2).ColIndex("VatValue")) = rsDummy("VatValue") & ""
    grd(2).TextMatrix(i, grd(2).ColIndex("VatBreakdownNote")) = rsDummy("VatBreakdownNote") & ""
End If

      grd(2).TextMatrix(i, grd(2).ColIndex("Ser")) = i
   grd(2).TextMatrix(i, grd(2).ColIndex("Transaction_ID")) = e.Transaction_ID
   grd(2).TextMatrix(i, grd(2).ColIndex("DefaultInvoicetype")) = e.Invoicetype

  grd(2).TextMatrix(i, grd(2).ColIndex("id")) = e.ID
  grd(2).TextMatrix(i, grd(2).ColIndex("ID700")) = e.Id700
  grd(2).TextMatrix(i, grd(2).ColIndex("chkTaxExempt")) = e.chkTaxExempt
  
  
     grd(2).TextMatrix(i, grd(2).ColIndex("IssueDate")) = e.IssueDate
      grd(2).TextMatrix(i, grd(2).ColIndex("IssueTim")) = e.IssueTim
     
       grd(2).TextMatrix(i, grd(2).ColIndex("InvoiceTypeCodeID")) = e.InvoiceTypeCodeID
  grd(2).TextMatrix(i, grd(2).ColIndex("InvoiceTypeCodename")) = e.InvoiceTypeCodename
 grd(2).TextMatrix(i, grd(2).ColIndex("DocumentCurrencyCode")) = e.DocumentCurrencyCode
 
  grd(2).TextMatrix(i, grd(2).ColIndex("TaxCurrencyCode")) = e.TaxCurrencyCode
   grd(2).TextMatrix(i, grd(2).ColIndex("InvoiceDocumentReferenceID")) = e.InvoiceDocumentReferenceID
    grd(2).TextMatrix(i, grd(2).ColIndex("AdditionalDocumentReferenceICVUUID")) = e.AdditionalDocumentReferenceICVUUID
     grd(2).TextMatrix(i, grd(2).ColIndex("ActualDeliveryDate")) = e.ActualDeliveryDate
      grd(2).TextMatrix(i, grd(2).ColIndex("LatestDeliveryDate")) = e.LatestDeliveryDate
      
       grd(2).TextMatrix(i, grd(2).ColIndex("PaymentMeansCode")) = e.PaymentMeansCode
        grd(2).TextMatrix(i, grd(2).ColIndex("InstructionNote")) = e.InstructionNote
        grd(2).TextMatrix(i, grd(2).ColIndex("PayeeFinancialAccount")) = e.PayeeFinancialAccount
         grd(2).TextMatrix(i, grd(2).ColIndex("paymentnote")) = e.paymentnote
           grd(2).TextMatrix(i, grd(2).ColIndex("Identificationid")) = e.Identificationid
        
        grd(2).TextMatrix(i, grd(2).ColIndex("Id700")) = IIf(IsNull(rsDummy("Id700").value), "", rsDummy("Id700").value)
         grd(2).TextMatrix(i, grd(2).ColIndex("chkTaxExempt")) = e.chkTaxExempt
  
        grd(2).TextMatrix(i, grd(2).ColIndex("schemeID")) = e.schemeID
        grd(2).TextMatrix(i, grd(2).ColIndex("StreetName")) = e.StreetName
        grd(2).TextMatrix(i, grd(2).ColIndex("AdditionalStreetName")) = e.AdditionalStreetName
        grd(2).TextMatrix(i, grd(2).ColIndex("BuildingNumber")) = e.BuildingNumber
        grd(2).TextMatrix(i, grd(2).ColIndex("PlotIdentification")) = e.PlotIdentification
        
         grd(2).TextMatrix(i, grd(2).ColIndex("CityName")) = e.CityName
        grd(2).TextMatrix(i, grd(2).ColIndex("PostalZone")) = e.PostalZone
        grd(2).TextMatrix(i, grd(2).ColIndex("CountrySubentity")) = e.CountrySubentity
        grd(2).TextMatrix(i, grd(2).ColIndex("CitySubdivisionName")) = e.CitySubdivisionName
        grd(2).TextMatrix(i, grd(2).ColIndex("IdentificationCode")) = e.IdentificationCode
        
   
  grd(2).TextMatrix(i, grd(2).ColIndex("RegistrationName")) = e.RegistrationName
    grd(2).TextMatrix(i, grd(2).ColIndex("CompanyID")) = e.CompanyID
 grd(2).TextMatrix(i, grd(2).ColIndex("allowancechargeAmount")) = e.allowancechargeAmount
    grd(2).TextMatrix(i, grd(2).ColIndex("AllowanceChargeReason")) = e.AllowanceChargeReason
    grd(2).TextMatrix(i, grd(2).ColIndex("TaxCategoryID")) = e.TaxCategoryID
   
    grd(2).TextMatrix(i, grd(2).ColIndex("TaxCategoryPercent")) = e.TaxCategoryPercent
    grd(2).TextMatrix(i, grd(2).ColIndex("PayableAmount")) = e.PayableAmount
    grd(2).TextMatrix(i, grd(2).ColIndex("PrepaidAmount")) = e.PrepaidAmount


   grd(2).TextMatrix(i, grd(2).ColIndex("InvoiceHash")) = e.InvoiceHash
   
    Dim strFullText As String
                    Dim strShortText As String
                    strFullText = e.SingedXML
                    ' strFullText = «·‰’ «·√’·Ì („Õ ÊÏ XML √Ê √Ì ‰’ ÿÊÌ·)
                    
                    If Len(strFullText) > 2000 Then
                    strShortText = left(strFullText, 2000)
                    Else
                    strShortText = strFullText
                    End If
                    
   grd(2).TextMatrix(i, grd(2).ColIndex("SingedXML")) = strShortText
   
                       strFullText = e.EncodedInvoice
                    ' strFullText = «·‰’ «·√’·Ì („Õ ÊÏ XML √Ê √Ì ‰’ ÿÊÌ·)
                    
                    If Len(strFullText) > 2000 Then
                    strShortText = left(strFullText, 2000)
                    Else
                    strShortText = strFullText
                    End If
   
  grd(2).TextMatrix(i, grd(2).ColIndex("EncodedInvoice")) = strShortText
   grd(2).TextMatrix(i, grd(2).ColIndex("UUID")) = e.UUID
   grd(2).TextMatrix(i, grd(2).ColIndex("QRCode")) = e.QRCode
  grd(2).TextMatrix(i, grd(2).ColIndex("PIH")) = e.PIH
   grd(2).TextMatrix(i, grd(2).ColIndex("SingedXMLFileName")) = e.SingedXMLFileName
grd(2).TextMatrix(i, grd(2).ColIndex("QrCodeDataPath")) = e.QrCodeDataPath
grd(2).TextMatrix(i, grd(2).ColIndex("DocType")) = e.docType2

   
   
   
  ' e.generateInvoice
     
       
        i = i + 1
        grd(2).rows = grd(2).rows + 1
        rsDummy.MoveNext
    Loop

End Sub

Private Sub Command3_Click()

    Dim objExcel As Object
    Dim objWorkbook As Object
    Dim objWorksheet As Object
    Dim i As Integer
    Dim j As Integer

    ' ≈‰‘«¡ ﬂ«∆‰ Excel ÃœÌœ
    Set objExcel = CreateObject("Excel.Application")
    Set objWorkbook = objExcel.Workbooks.Add
    Set objWorksheet = objWorkbook.Sheets(1)

    '  ’œÌ— ⁄‰«ÊÌ‰ «·√⁄„œ…
    For j = 0 To grd(0).Cols - 1
        objWorksheet.cells(1, j + 1).value = grd(0).ColKey(j)
    Next j

    '  ’œÌ— «·»Ì«‰« 
'    For i = 0 To grd(0).rows - 1
'        For j = 0 To grd(0).Cols - 1
'            objWorksheet.Cells(i + 2, j + 1).value = grd(0).Columns(j).CellText(i)
'        Next j
'    Next i

    '  ‰”Ìﬁ «·⁄„Êœ
    objWorksheet.Columns.AutoFit

    ' ⁄—÷ Excel Ê≈ÿ·«ﬁÂ
    objExcel.Visible = True

    '  ‰ŸÌ› «·ﬂ«∆‰« 
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing


End Sub


Private Function CleanFileName(ByVal FileName As String) As String
    Dim BadChars As Variant
    Dim i As Integer

    ' —„Ê“ €Ì— „”„ÊÕ… ›Ì «”„«¡ «·„·›« 
    BadChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    ' «” »œ«· «·—„Ê“
    For i = LBound(BadChars) To UBound(BadChars)
        FileName = Replace(FileName, BadChars(i), "-")
    Next i
    
    ' „„ﬂ‰ ﬂ„«‰ ‰‘Ì· «·„”«›«  ·Ê  Õ»
    FileName = Replace(FileName, " ", "_")

    CleanFileName = FileName
End Function

Private Sub Command4_Click(Index As Integer)
Dim reporttitle As String

    If Index = 0 Then
        reporttitle = "Invoices_Not_Submitted" & "_From_" & Format(FromDate.value, "yyyy-mm-dd") & "_To_" & Format(ToDate.value, "yyyy-mm-dd")
    Else
        reporttitle = "Invoices_Submitted" & "_From_" & Format(FromDate10.value, "yyyy-mm-dd") & "_To_" & Format(ToDate10.value, "yyyy-mm-dd")
    End If

' ‰Ÿ› «·«”„ ﬁ»· «·«—”«·
reporttitle = CleanFileName(reporttitle)

' «” œ⁄«¡ «· ’œÌ—
If Index = 0 Then
    ExportToExcel Me, grd(0), "›Ê« Ì— ·„  —›⁄", , reporttitle
Else
    ExportToExcel Me, grd(2), "›Ê« Ì—  „ —›⁄Â«", , reporttitle
End If

End Sub

Private Sub DBCboClientName_Click(Area As Integer)
Dim fullcode As String

 GetCustomersDetail val(DBCboClientName.BoundText), , fullcode, 1
    TxtSearchCode.Text = fullcode
    
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub

Sub FillPayment()
Dim i As Integer
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
With CboPayMentType
.Clear
.AddItem "‰ﬁœÌ"
End With
sql = "SELECT        PaymentID, PaymentName, PaymentNamee"
sql = sql & " From dbo.TblPaymentType"
sql = sql & " order by PaymentID "
Rs3.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
For i = 1 To Rs3.RecordCount
With CboPayMentType
.AddItem IIf(IsNull(Rs3("PaymentName").value), "", Rs3("PaymentName").value)
End With
Rs3.MoveNext
Next i
End If
Rs3.Close
End Sub

Private Sub grd_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
Dim mPath As String
With grd(Index)

           mPath = App.path
                                 If mId(mPath, Len(Trim(mPath)), 1) = "\" Then
                                     mPath = left(mPath, Len(Trim(mPath)) - 1)
                                 End If
                                 
        Select Case .ColKey(Col)
Case "View", "Id", "InvoiceID"
        
            
            If val(grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodeID"))) = 388 Then
                frmsalebill.TxtModFlg = "R"
                Unload frmsalebill
                frmsalebill.show
                frmsalebill.XPBtnMove_Click (2)
                frmsalebill.Retrive val(grd(Index).TextMatrix(.Row, grd(Index).ColIndex("Transaction_ID")))
                frmsalebill.TxtModFlg = "R"
            ElseIf val(grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodeID"))) = 381 Then
                FrmReturnSalling.TxtModFlg = "R"
                Unload FrmReturnSalling
                FrmReturnSalling.show
                FrmReturnSalling.XPBtnMove_Click (2)
                FrmReturnSalling.Retrive val(grd(Index).TextMatrix(.Row, grd(Index).ColIndex("Transaction_ID")))
                FrmReturnSalling.TxtModFlg = "R"

            End If
            Case "viewFile"
            If val(grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodeID"))) = "388" And (grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodename"))) = "0100000" Then
             ShellExecute 0&, vbNullString, mPath & "\Invoices\Signed\Standard\Invoices\" & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("SingedXMLFileName")), vbNullString, vbNullString, vbNormalFocus
             ElseIf val(grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodeID"))) = "388" And (grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodename"))) = "0200000" Then
             ShellExecute 0&, vbNullString, mPath & "\Invoices\Signed\Simplified\Invoices\" & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("SingedXMLFileName")), vbNullString, vbNullString, vbNormalFocus
           ElseIf val(grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodeID"))) = "383" And (grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodename"))) = "0100000" Then
             ShellExecute 0&, vbNullString, mPath & "\Invoices\Signed\Standard\Debit\" & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("SingedXMLFileName")), vbNullString, vbNullString, vbNormalFocus
                  ElseIf val(grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodeID"))) = "383" And (grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodename"))) = "0200000" Then
             ShellExecute 0&, vbNullString, mPath & "\Invoices\Signed\Simplified\Debit\" & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("SingedXMLFileName")), vbNullString, vbNullString, vbNormalFocus
                  ElseIf val(grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodeID"))) = "381" And (grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodename"))) = "0100000" Then
             ShellExecute 0&, vbNullString, mPath & "\Invoices\Signed\Standard\Credit\" & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("SingedXMLFileName")), vbNullString, vbNullString, vbNormalFocus
                   ElseIf val(grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodeID"))) = "381" And (grd(Index).TextMatrix(.Row, grd(Index).ColIndex("InvoiceTypeCodename"))) = "0200000" Then
             ShellExecute 0&, vbNullString, mPath & "\Invoices\Signed\Simplified\Credit\" & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("SingedXMLFileName")), vbNullString, vbNullString, vbNormalFocus
                             
                 End If
                            
                           
                
                Case "ErrorMessage"
                MsgBox "invoice no:" & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("id")) & CHR(13) & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("ErrorMessage"))
      Case "warrningmessage"
                MsgBox "invoice no:" & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("id")) & CHR(13) & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("warrningmessage"))
 
 
 
               Case "viewQRCode"
               Picture1.Visible = True
     If grd(Index).TextMatrix(.Row, grd(Index).ColIndex("QrCodeDataPath")) <> "" Then
        Picture1.Picture = LoadPicture(grd(Index).TextMatrix(.Row, grd(Index).ColIndex("QrCodeDataPath")))
      
    Else
      
      Set Picture1.Picture = Nothing
      
    End If
    
    
                 Case "ViewError"
                 MsgBox grd(Index).TextMatrix(.Row, grd(Index).ColIndex("ErrorMessage"))
               
        End Select

    End With

 
End Sub

Private Sub grd_Click(Index As Integer)
              With grd(Index)
        Select Case .ColKey(grd(Index).Col)

          
      Case "warrningmessage"
                MsgBox "invoice no:" & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("id")) & CHR(13) & grd(Index).TextMatrix(.Row, grd(Index).ColIndex("warrningmessage"))
 
 End Select
 End With
End Sub

Private Sub Picture1_Click()
Picture1.Visible = False
End Sub

Private Sub ConectionFirst()
    On Error GoTo ConnErr

    If Not POSConnection Is Nothing Then
        If POSConnection.State = adStateOpen Then
            POSConnection.Close
        End If
        Set POSConnection = Nothing
    End If
    Dim SysSQLServerUserpassword2 As String
SysSQLServerUserpassword2 = "Admin.com"
'SysSQLServerUserpassword2 = "Admin@123"

    Set POSConnection = New ADODB.Connection
    With POSConnection
        .CommandTimeout = 5000
        .CursorLocation = adUseClient
        .ConnectionTimeout = 5000
        If SysSQLServerType = 1 Then
            .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword2 & _
                                ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                                ";Initial Catalog=" & SystemOptions.DbNameW & _
                                ";Data Source=" & SystemOptions.ServerNameW
        ElseIf SysSQLServerType = 2 Then
            If SysSQLServerTypeTechnical = "0" Then
                .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;" & _
                                    "Initial Catalog=" & SystemOptions.DbNameW & _
                                    ";Data Source=" & SystemOptions.ServerNameW
            Else
                .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SysSQLServerUserpassword2 & _
                                    ";Persist Security Info=True;User ID=" & SysSQLServerUserId & _
                                    ";Initial Catalog=" & SystemOptions.DbNameW & _
                                    ";Data Source=" & SystemOptions.ServerNameW
            End If
        End If
        .Open
    End With
    Exit Sub

ConnErr:
    MsgBox "›‘· «·« ’«· »ﬁ«⁄œ… «·»Ì«‰« !" & vbCrLf & Err.Description, vbCritical, "Œÿ√ ›Ì «·« ’«·"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

If Text3.Text = "01271316739" Then
    Dim s As String
    ' «· √ﬂœ „‰ √‰ «·ﬁÌ„ «·√’·Ì… €Ì— ›«—€… ﬁ»· «” —Ã«⁄Â«
    s = "UPDATE TblOptions SET Privatekey = Privatekey2, PublickeycertPem = PublickeycertPem2, SecretKey = SecretKey2 " & _
        "WHERE Privatekey2 IS NOT NULL AND PublickeycertPem2 IS NOT NULL AND SecretKey2 IS NOT NULL"
    Cn.Execute s
    
    '  ‰ŸÌ› «·ÕﬁÊ· «·«Õ Ì«ÿÌ… »⁄œ «·«” —Ã«⁄
    s = "UPDATE TblOptions SET Privatekey2 = NULL, PublickeycertPem2 = NULL, SecretKey2 = NULL"
    Cn.Execute s
    
    ' ‰›” «·⁄„·Ì… ·ÃœÊ· TblBranchesData
    s = "UPDATE TblBranchesData SET Privatekey = Privatekey2, PublickeycertPem = PublickeycertPem2, SecretKey = SecretKey2 " & _
        "WHERE Privatekey2 IS NOT NULL AND PublickeycertPem2 IS NOT NULL AND SecretKey2 IS NOT NULL"
    Cn.Execute s
    
    '  ‰ŸÌ› «·ÕﬁÊ· «·«Õ Ì«ÿÌ… »⁄œ «·«” —Ã«⁄
    s = "UPDATE TblBranchesData SET Privatekey2 = NULL, PublickeycertPem2 = NULL, SecretKey2 = NULL"
    Cn.Execute s
    MsgBox "Password error"
End If

If Text3.Text = SystemOptions.BigUserPw2 Then
Command1.Visible = True
Else
Command1.Visible = False
End If
End Sub

Private Sub TxtItemCode_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim Msg As String
    Dim StrSQL As String
    Dim rs As ADODB.Recordset

            If KeyCode = vbKeyReturn Then
                If Trim(Me.TxtItemCode(Index).Text) = "" Then Exit Sub
                StrSQL = "Select ItemID From TblItems Where ItemCode='" & Trim(Me.TxtItemCode(Index).Text) & "'"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    DCboItemsName.BoundText = rs("ItemID").value
                Else
                    Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
            End If



 If KeyCode = vbKeyF3 Then
            Load FrmItemSearch
            FrmItemSearch.RetrunType = 1
            Set FrmItemSearch.DcboItems = Me.DCboItemsName
            FrmItemSearch.show vbModal

End If

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub



Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    FillPayment
    ToDate.value = Date
    FromDate.value = Date
    
        ToDate10.value = Date
    FromDate10.value = Date
    
    DTP_Date.value = Date
    'fillmycompanydata
    
    Set Dcombos = New ClsDataCombos
      Dcombos.GetStores Me.DCboStoreName2
      Dcombos.GetBranches Me.DcbBranch
      Dcombos.GetBranches Me.DcBranches(3)
      Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
      Dcombos.GetItemsSizes Me.DcbSize
      Dcombos.GetItemsColors Me.DcbColor
      Dcombos.GetItemsNames Me.DCboItemsName
      Dcombos.GetItemsNames Me.DCboItemsName2
      Dcombos.GetStores Me.DCboStoreName
      Dcombos.GetItemSGroups Me.DCboGroup1, False
      Dcombos.GetSalesRepData Me.DcbEmp
      Dcombos.GetBranches Me.DcbBranch2
      Dcombos.GetSection Me.DCRegionID
      Dcombos.GetSection Me.DCRegionID2
     If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  select id,name from tblActivitesType   "
    Else
        StrSQL = "  select id,namee from tblActivitesType   "
    End If
    fill_combo DCActivity, StrSQL
    fill_combo DCActivity2, StrSQL

    fill_combo DcBranches(2), StrSQL

  C1Tab1.TabVisible(3) = False
  C1Tab1.TabVisible(4) = False
  C1Tab1.CurrTab = 1
  
  If SystemOptions.IsBlue Then
        If mIndex = 4 Then
          C1Tab1.TabVisible(4) = True
          Me.Caption = "«⁄„«— «·«’‰«›"
          Label5(0).Caption = "«⁄„«— «·«’‰«›"
          C1Tab1.CurrTab = 4
        End If
End If
If mIndex = 3 Then
    
        
     C1Tab1.TabVisible(0) = False
     C1Tab1.TabVisible(1) = False
     C1Tab1.TabVisible(2) = False
     C1Tab1.TabVisible(3) = False
     C1Tab1.TabVisible(4) = False
  
  
     C1Tab1.TabVisible(mIndex) = True
     C1Tab1.CurrTab = mIndex
     Label5(0).Caption = " "
     Me.Caption = " "
End If
If mIndex <> 3 Then
    Me.Width = Me.Width - Frame4.left
    Me.Height = Me.Height - 1800
End If

If mIndex <> 3 Then
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
End If
    DTPicker2.value = Date
    DTPicker1.value = Date
    DTPicker2.value = ""
    DTPicker1.value = ""
DtpDateFrom.value = ""
DtpDateTo.value = ""
DtpDateFrom2.value = Date
DtpDateTo2.value = Date
    If mIndex = 3 Then
        Resize_Form Me, TransactionSize
        'Me.MDIChild = False
        Me.WindowState = vbMaximized
        Me.show
    Else
    Resize_Form Me
    
    End If
              
    
 
        

End Sub


Private Sub Form_Unload(Cancel As Integer)
mIndex = 0
    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Sub GetDataTrans3()
Dim sql As String
Dim BrnchesReg As String
Dim BrnchAct As String
    BrnchesReg = BranchRegion(val(DCRegionID2.BoundText))
    BrnchAct = BrcnhActivityType(val(DCActivity.BoundText))
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset


sql = " SELECT        MONTH(Transactions.Transaction_Date) AS mMonth,Year(Transactions.Transaction_Date) as mYear, TblEmployee.Commission, Transaction_Details.EmpID4, TblEmployee.Emp_Name,"
sql = sql & " TblEmpDepartmentsDet.Name DeptName2,TblEmpDepartments.DepartmentName,"
sql = sql & " sum(Transaction_Details.showPrice * Transaction_Details.Quantity )  Total ,TblEmployee.Emp_Code,"
sql = sql & "                          SUM (DOUBLE_ENTREY_VOUCHERS.value) as Salary"
sql = sql & " FROM            Notes INNER JOIN"
sql = sql & "                          DOUBLE_ENTREY_VOUCHERS ON Notes.NoteID = DOUBLE_ENTREY_VOUCHERS.Notes_ID"
                        
sql = sql & "                       RIGHT OUTER JOIN"
sql = sql & "                          ACCOUNTS ON DOUBLE_ENTREY_VOUCHERS.Account_Code = ACCOUNTS.Account_Code FULL OUTER JOIN"
sql = sql & "                          TblEmployee ON ACCOUNTS.Account_Code = TblEmployee.Account_code1 FULL OUTER JOIN"
sql = sql & "                          Transaction_Details INNER JOIN"
sql = sql & "                          Transactions ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID ON TblEmployee.Emp_ID = Transaction_Details.EmpID4"
sql = sql & "                           and MONTH(Transactions.Transaction_Date) = MONTH(Notes.NoteDate)"
sql = sql & "                           and Year(Transactions.Transaction_Date) = Year(Notes.NoteDate)"
sql = sql & "                           Left outer join TblEmpDepartmentsDet On TblEmployee.DeptID2 = TblEmpDepartmentsDet.DeparmentID"
sql = sql & "                           Left outer join TblEmpDepartments On TblEmployee.DepartmentID = TblEmpDepartments.DeparmentID"
sql = sql & " Where (transactions.Transaction_Type = 21) And (IsNull(Transaction_Details.EmpID4, 0) <> 0)"


If val(DcbBranch2.BoundText) <> 0 Then
sql = sql & " and dbo.Notes.branch_no =" & val(DcbBranch2.BoundText) & ""
Else
sql = sql & " AND      dbo.Notes.branch_no  in(" & Current_branchSql & ")"
End If
      If Not IsNull(Me.DTPicker1.value) Then
                   sql = sql & " AND Month(dbo.transactions.Transaction_Date) >=" & Month(SQLDate(Me.DTPicker1.value, True)) & ""
      End If
       If Not IsNull(Me.DTPicker2.value) Then
                   
                   sql = sql & " AND Month(dbo.transactions.Transaction_Date) <=" & Month(SQLDate(Me.DTPicker2.value, True)) & ""
      End If
sql = sql & " Group By"
sql = sql & " Month (transactions.Transaction_Date),Year(Transactions.Transaction_Date), TblEmployee.commission, Transaction_Details.EmpID4, TblEmployee.emp_Name"
sql = sql & " ,TblEmployee.Emp_Code,TblEmpDepartmentsDet.Name ,TblEmpDepartments.DepartmentName"
 
    
print_report sql, 50
End Sub

Sub GetDataTrans2()
Dim sql As String
Dim BrnchesReg As String
Dim BrnchAct As String
    BrnchesReg = BranchRegion(val(DCRegionID2.BoundText))
    BrnchAct = BrcnhActivityType(val(DCActivity.BoundText))
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     TradingContractID, dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value, dbo.Notes.CashingType, "
sql = sql & "                     dbo.Notes.CusID, dbo.TblCustemers.Fullcode, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Notes.project_id, dbo.projects.Project_name,"
sql = sql & "                      dbo.projects.Project_account, dbo.projects.opening_balance_voucher_id, isnull(dbo.TblMultuPayment.PaymentID,0)as PaymentID, dbo.TblMultuPayment.[Value],"
sql = sql & "                      dbo.TblMultuPayment.CardNo, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Notes.branch_no, dbo.TblBranchesData.branch_name,"
sql = sql & "                      dbo.TblBranchesData.branch_namee, dbo.Notes.NoteCashingType, dbo.Notes.EmployeeID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS Expr1,"
sql = sql & "                      dbo.TblEmployee.Emp_Namee , dbo.Notes.AccountsCode, dbo.Accounts.account_name, dbo.Accounts.account_serial, dbo.Accounts.Account_NameEng"
sql = sql & " FROM         dbo.ACCOUNTS RIGHT OUTER JOIN"
sql = sql & "                      dbo.Notes ON dbo.ACCOUNTS.Account_Code = dbo.Notes.AccountsCode LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.Notes.EmployeeID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblPaymentType RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblMultuPayment ON dbo.TblPaymentType.PaymentID = dbo.TblMultuPayment.PaymentID ON dbo.Notes.NoteID = dbo.TblMultuPayment.NoteID AND ISNULL(TblMultuPayment.[Value],0) <> 0 LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.Notes.project_id = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID"
sql = sql & " Where (dbo.Notes.notetype = 4) "
If BrnchesReg <> "-1" Then
        sql = sql & " AND dbo.Notes.branch_no in( " & BrnchesReg & " )"
End If
If BrnchAct <> "-1" Then
        sql = sql & " AND dbo.Notes.branch_no in( " & BrnchAct & " )"
End If
If val(DcbBranch2.BoundText) <> 0 Then
sql = sql & " and dbo.Notes.branch_no =" & val(DcbBranch2.BoundText) & ""
Else
sql = sql & " AND      dbo.Notes.branch_no  in(" & Current_branchSql & ")"
End If
      If Not IsNull(Me.DTPicker1.value) Then
                   sql = sql & " AND dbo.Notes.NoteDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
      End If
       If Not IsNull(Me.DTPicker2.value) Then
                   sql = sql & " AND dbo.Notes.NoteDate <=" & SQLDate(Me.DTPicker2.value, True) & ""
      End If
    sql = sql & " order by dbo.TblMultuPayment.PaymentID"
print_report2 sql, 2
End Sub
Sub GetDataTrans()
Dim sql As String
Dim BrnchesReg As String
Dim BrnchAct As String

       BrnchesReg = BranchRegion(val(DCRegionID2.BoundText))
      BrnchAct = BrcnhActivityType(val(DCActivity.BoundText))

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.TransactionTypes.TransactionTypeName, "
sql = sql & "                      dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_HijriDate, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
sql = sql & "                      dbo.TblCustemers.Fullcode, dbo.Transactions.CusID, dbo.Transactions.NoteSerial1, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile,"
sql = sql & "                      dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile, dbo.TblSalesPayment.[Value],"
sql = sql & "                      dbo.TblSalesPayment.CardNo,"
sql = sql & "      PaymentName=Case "
sql = sql & "     When  Transactions.PaymentType=1   Then '«Ã·' "
 sql = sql & "    Else  "
 
sql = sql & "   ISNULL(dbo.TblPaymentType.PaymentName, N'‰ﬁœÌ')  "
sql = sql & "           END,"
sql = sql & "       dbo.TblPaymentType.PaymentNamee,"
'ISNULL(dbo.TblPaymentType.PaymentName, '????') AS PaymentName, dbo.TblPaymentType.PaymentNamee,"
sql = sql & "                      dbo.Transactions.PaymentType, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
sql = sql & "                      dbo.Transactions.Transaction_NetValue, dbo.Transactions.VAT, ISNULL(dbo.TblSalesPayment.PaymentID, 0) AS PaymentID"
sql = sql & "  FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
sql = sql & "                       dbo.Transactions ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId LEFT OUTER JOIN"
sql = sql & "                       dbo.TblPaymentType RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblSalesPayment ON dbo.TblPaymentType.PaymentID = dbo.TblSalesPayment.PaymentID ON"
sql = sql & "                       dbo.Transactions.Transaction_ID = dbo.TblSalesPayment.TransID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
If Optrans(2).value = True Then
sql = sql & "  WHERE  (dbo.Transactions.Transaction_Type = 21) AND (dbo.TblSalesPayment.[Value] <> 0 OR"
sql = sql & "                      dbo.TblSalesPayment.[Value] IS NULL)"
Else
sql = sql & "  WHERE  (dbo.Transactions.Transaction_Type = 21 OR"
sql = sql & "                      dbo.Transactions.Transaction_Type = 22) AND (dbo.TblSalesPayment.[Value] <> 0 OR"
sql = sql & "                      dbo.TblSalesPayment.[Value] IS NULL)"
End If
    If BrnchesReg <> "-1" Then
        sql = sql & " AND dbo.Transactions.BranchId in( " & BrnchesReg & " )"
    End If
        If BrnchAct <> "-1" Then
        sql = sql & " AND dbo.Transactions.BranchId in( " & BrnchAct & " )"
    End If
If val(DcbBranch2.BoundText) <> 0 Then
sql = sql & " AND dbo.Transactions.BranchId =" & val(DcbBranch2.BoundText) & ""
Else
sql = sql & " AND      dbo.Transactions.BranchId   in(" & Current_branchSql & ")"
End If
      If Not IsNull(Me.DTPicker1.value) Then
                   sql = sql & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DTPicker1.value, True) & ""
      End If
       If Not IsNull(Me.DTPicker2.value) Then
                   sql = sql & " AND dbo.Transactions.Transaction_Date<=" & SQLDate(Me.DTPicker2.value, True) & ""
      End If
      
      
    
      
            
      
    sql = sql & " order by dbo.Transactions.Transaction_Type, dbo.TblSalesPayment.PaymentID"
print_report2 sql, 1
End Sub
Public Sub GetData()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
If opt(0).value = True Or opt(1).value = True Or opt(2).value = True Or opt(5).value = True Or opt(6).value = True Or opt(9).value = True Or opt(10).value = True Or opt(13).value = True Then
reportid = 0
ElseIf opt(15).value = True Then
reportid = 15
Else
reportid = 1
End If

'StrSQL = " SELECT     TOP 100 PERCENT dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, SUM(dbo.ItemsDetails.[Count] * dbo.TransactionTypes.StockEffect) "
'StrSQL = StrSQL & "                       AS countsactual, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName AS [cclASS NAME], dbo.TblItemsColors.ColorName,"
'StrSQL = StrSQL & "                       dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.StoreID, dbo.Transactions.Transaction_Date,"
'StrSQL = StrSQL & "                       dbo.TblItems.ItemCode , dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.ItemsDetails.ItemID"
'StrSQL = StrSQL & "  FROM         dbo.ItemsDetails INNER JOIN"
'StrSQL = StrSQL & "                       dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
'StrSQL = StrSQL & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
'StrSQL = StrSQL & "                       dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                       dbo.TblItems ON dbo.ItemsDetails.ItemId = dbo.TblItems.ItemID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                       dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                       dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                       dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
'StrSQL = StrSQL & "                       dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
'StrSQL = StrSQL & " where 1=1"


'StrSQL = "SELECT     dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, SUM(dbo.ItemsDetails.[Count] * dbo.TransactionTypes.StockEffect) "
'StrSQL = StrSQL & " AS countsactual, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName AS [cclASS NAME], dbo.TblItemsColors.ColorName,"
'StrSQL = StrSQL & " dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.StoreID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
'StrSQL = StrSQL & " dbo.TblItems.ItemNamee , dbo.ItemsDetails.ItemID"
'StrSQL = StrSQL & " FROM         dbo.ItemsDetails INNER JOIN"
'StrSQL = StrSQL & " dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
'StrSQL = StrSQL & "   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
'StrSQL = StrSQL & "   dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
'StrSQL = StrSQL & "   dbo.TblItems ON dbo.ItemsDetails.ItemId = dbo.TblItems.ItemID LEFT OUTER JOIN"
'StrSQL = StrSQL & "   dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
'StrSQL = StrSQL & "   dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
'StrSQL = StrSQL & "   dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
'StrSQL = StrSQL & "    dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
 
 StrSQL = "  SELECT     SUM(dbo.ItemsDetails.[Count] * dbo.ItemsDetails.EffectN) AS countsactual, "
StrSQL = StrSQL & " dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName AS [cclASS NAME], dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & " dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.StoreID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
StrSQL = StrSQL & " dbo.ItemsDetails.ItemID , dbo.Groups.GroupName, dbo.Groups.GroupNamee"
StrSQL = StrSQL & "  FROM         dbo.Groups INNER JOIN"
StrSQL = StrSQL & "  dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID RIGHT OUTER JOIN"
StrSQL = StrSQL & "  dbo.ItemsDetails INNER JOIN"
StrSQL = StrSQL & "  dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblItems.ItemID = dbo.ItemsDetails.ItemId LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
 

StrSQL = StrSQL & "   Where (1 = 1)"

 
 If reportid = 1 Then
 StrSQL = "SELECT       dbo.Transactions.NoteSerial1, dbo.TransactionTypes.Transaction_Type, dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, "
StrSQL = StrSQL & "   dbo.ItemsDetails.ParrtNoCode, SUM(dbo.ItemsDetails.[Count] * dbo.ItemsDetails.EffectN) AS countsactual,"
StrSQL = StrSQL & " dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName AS [cclASS NAME], dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & " dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.StoreID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
StrSQL = StrSQL & " dbo.ItemsDetails.ItemID , dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID"
StrSQL = StrSQL & " FROM         dbo.Groups INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TransactionTypes INNER JOIN"
StrSQL = StrSQL & "                      dbo.Transactions ON dbo.TransactionTypes.Transaction_Type = dbo.Transactions.Transaction_Type INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.ItemsDetails ON dbo.Transactions.Transaction_ID = dbo.ItemsDetails.Transaction_ID ON dbo.TblItems.ItemID = dbo.ItemsDetails.ItemId LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
StrSQL = StrSQL & "  Where (1 = 1)  "

 End If
 '''''''''''''''''''
 
  If ItemDetailedCode.Text <> "" Then
     StrSQL = StrSQL & " AND dbo.ItemsDetails.ItemDetailedCode like '%" & Me.ItemDetailedCode.Text & "%'"
    End If
    
     If ParrtNoCode.Text <> "" Then
     StrSQL = StrSQL & " AND dbo.ItemsDetails.ParrtNoCode like '%" & Me.ParrtNoCode.Text & "%'"
    End If
    
    
    
    
    
If Me.DCboStoreName.Text <> "" And val(DCboStoreName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.Transactions.StoreID = " & val(Me.DCboStoreName.BoundText)

End If

If Me.DCboItemsName.Text <> "" And val(DCboItemsName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.ItemId = " & val(Me.DCboItemsName.BoundText)

End If
If Me.DcbColor.Text <> "" And val(DcbColor.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.ColorID = " & val(Me.DcbColor.BoundText)

End If
If Me.DcbSize.Text <> "" And val(DcbSize.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.SizeID = " & val(Me.DcbSize.BoundText)

End If

If Me.DCboGroup1.Text <> "" And val(DCboGroup1.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND    dbo.TblItems.GroupID= " & val(Me.DCboGroup1.BoundText)

End If

If opt(0).value = True Then

ElseIf opt(1).value = True Or opt(3).value = True Then
 StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =21"
 
 ElseIf opt(1).value = True Or opt(3).value = True Then
 StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =21"
 
 ElseIf opt(2).value = True Or opt(4).value = True Then
 StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =22"
 ElseIf opt(5).value = True Or opt(7).value = True Then
 StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =9"
 ElseIf opt(6).value = True Or opt(8).value = True Then
 StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =5"
 
 ElseIf opt(9).value = True Or opt(11).value = True Then
 StrSQL = StrSQL & " AND  ( dbo.Transactions.Transaction_Type =9  or dbo.Transactions.Transaction_Type =21 )"
 
 
 ElseIf opt(10).value = True Or opt(12).value = True Then
 StrSQL = StrSQL & " AND  ( dbo.Transactions.Transaction_Type =5 or  dbo.Transactions.Transaction_Type =22 )"
  ElseIf opt(13).value = True Or opt(14).value = True Then
  StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =3"
End If






 If DBCboClientName.BoundText <> "" And DBCboClientName.Text <> "" Then
                   StrSQL = StrSQL & " AND dbo.Transactions.CusID =" & val(DBCboClientName.BoundText)
      End If
 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
  If reportid = 0 Then
 StrSQL = StrSQL & "  GROUP BY  dbo.ItemsDetails.ColorID, dbo.ItemsDetails.ClassId, dbo.Transactions.StoreID, "
StrSQL = StrSQL & "  dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & "  dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.ItemsDetails.ItemId,"
StrSQL = StrSQL & "  dbo.Groups.GroupName , dbo.Groups.GroupNamee"
StrSQL = StrSQL & "  ORDER BY dbo.TblItems.ItemCode"
Else
StrSQL = StrSQL & "   GROUP BY dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.ClassId, dbo.Transactions.StoreID, "
StrSQL = StrSQL & "  dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & " dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.ItemsDetails.ItemId,"
StrSQL = StrSQL & " dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblItems.GroupID, dbo.TransactionTypes.TransactionTypeName,"
StrSQL = StrSQL & " dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID,"
StrSQL = StrSQL & " dbo.TransactionTypes.Transaction_Type,  dbo.Transactions.NoteSerial1"
 
StrSQL = StrSQL & "  ORDER BY dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID"

End If

If reportid = 15 Then
StrSQL = "SELECT       SUM(dbo.ItemsDetails.[Count] * dbo.ItemsDetails.EffectN) AS countsactual, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.SizeID, "

If val(DCboStoreName.BoundText) <> 0 Then
'strSQL = strSQL & "   dbo.ItemsDetails.ItemId, dbo.GardTransactionDetails1(" & SQLDate(Me.DtpDateFrom.value, True) & ", dbo.ItemsDetails.ColorID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ItemId,  " & val(DCboStoreName.BoundText) & ") AS QtyAvilable, "

 StrSQL = StrSQL & "   dbo.ItemsDetails.ItemId, dbo.GardTransactionDetails1('" & SQLDate(Date, False) & "', dbo.ItemsDetails.ColorID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ItemId, " & val(DCboStoreName.BoundText) & ") AS QtyAvilable, "
Else
 
StrSQL = StrSQL & "   dbo.ItemsDetails.ItemId, dbo.GardTransactionDetails1('" & SQLDate(Date, False) & "', dbo.ItemsDetails.ColorID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ItemId, Null) AS QtyAvilable, "
End If

StrSQL = StrSQL & "                        dbo.TblItems.ItemNamee , dbo.TblItems.ItemName, dbo.TblItemsSizes.sizename, dbo.TblItemsColors.colorname, dbo.TblItems.fullcode"
StrSQL = StrSQL & "  FROM         dbo.Groups INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.ItemsDetails INNER JOIN"
StrSQL = StrSQL & "                        dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
 StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblItems.ItemID = dbo.ItemsDetails.ItemId LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
StrSQL = StrSQL & "  WHERE     (1 = 1)   "
'StrSQL = StrSQL & "   (dbo.Transactions.StoreID = 2) "
'StrSQL = StrSQL & "   AND (dbo.ItemsDetails.ItemId = 70)"
'StrSQL = StrSQL & "   AND (dbo.Transactions.Transaction_Date >= '01-Oct-2016')  "
'StrSQL = StrSQL & "    And                    (dbo.Transactions.Transaction_Date <= '01-Oct-2016') "
StrSQL = StrSQL & "   AND (dbo.Transactions.Transaction_Type = 21)"

  If ItemDetailedCode.Text <> "" Then
     StrSQL = StrSQL & " AND dbo.ItemsDetails.ItemDetailedCode like '%" & Me.ItemDetailedCode.Text & "%'"
    End If
    
     If ParrtNoCode.Text <> "" Then
     StrSQL = StrSQL & " AND dbo.ItemsDetails.ParrtNoCode like '%" & Me.ParrtNoCode.Text & "%'"
    End If
    
    
    
    
    
If Me.DCboStoreName.Text <> "" And val(DCboStoreName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.Transactions.StoreID = " & val(Me.DCboStoreName.BoundText)

End If

If Me.DCboItemsName.Text <> "" And val(DCboItemsName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.ItemId = " & val(Me.DCboItemsName.BoundText)

End If
If Me.DcbColor.Text <> "" And val(DcbColor.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.ColorID = " & val(Me.DcbColor.BoundText)

End If
If Me.DcbSize.Text <> "" And val(DcbSize.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.SizeID = " & val(Me.DcbSize.BoundText)

End If

If Me.DCboGroup1.Text <> "" And val(DCboGroup1.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND    dbo.TblItems.GroupID= " & val(Me.DCboGroup1.BoundText)

End If


 If DBCboClientName.BoundText <> "" And DBCboClientName.Text <> "" Then
                   StrSQL = StrSQL & " AND dbo.Transactions.CusID =" & val(DBCboClientName.BoundText)
      End If
 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
      
      
StrSQL = StrSQL & "  GROUP BY dbo.ItemsDetails.ColorID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ItemId, dbo.TblItems.ItemNamee, dbo.TblItems.ItemName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & "                        dbo.TblItemsColors.colorname , dbo.TblItems.fullcode"
StrSQL = StrSQL & "   ORDER BY dbo.TblItemsColors.ColorName"


End If

    
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
    End If

End Sub
Public Sub GetDataNetwork()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim BrnchesReg As String
    Dim BrnchAct As String
        BrnchesReg = BranchRegion(val(DCRegionID.BoundText))
        BrnchAct = BrcnhActivityType(val(DCActivity2.BoundText))
        
        If optNetWork(3).value = True Or optNetWork(4).value = True Then
 '       StrSQL = "  SELECT     dbo.Transactions.last_changed,   dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue, dbo.Transactions.Transaction_ID, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, "
 '       StrSQL = StrSQL & "                                       dbo.Transactions.noteserial1, { fn HOUR(dbo.Transactions.last_changed) } AS hourx"
 '       StrSQL = StrSQL & "                 FROM            dbo.Transactions INNER JOIN"
 '       StrSQL = StrSQL & "                                          dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
 '       StrSQL = StrSQL & "                 Where (dbo.transactions.Transaction_Type = 21)"

StrSQL = "SELECT     CAST(last_changed AS TIME) Time2,dbo.Transactions.last_changed, dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue, "
     StrSQL = StrSQL & "                               dbo.Transactions.Transaction_ID, dbo.Transactions.NoteSerial1, { fn HOUR(dbo.Transactions.last_changed) } AS hourx, dbo.TblBranchesData.branch_id,"
     StrSQL = StrSQL & "                               dbo.TblEmployee.Emp_Name, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.Emp_Namee, dbo.TblStore.StoreName,"
     StrSQL = StrSQL & "                               dbo.TblStore.storenamee"
     StrSQL = StrSQL & "         FROM         dbo.Transactions INNER JOIN"
     StrSQL = StrSQL & "                               dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
     StrSQL = StrSQL & "                               dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID INNER JOIN"
     StrSQL = StrSQL & "                               dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
     StrSQL = StrSQL & "         WHERE     (dbo.Transactions.Transaction_Type = 21) "
                            
    If Me.DcbEmp.Text <> "" And val(DcbEmp.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND    Emp_ID= " & val(Me.DcbEmp.BoundText)
    End If
    
                            If Me.DCboStoreName2.Text <> "" And val(DCboStoreName2.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   Transactions.StoreID = " & val(Me.DCboStoreName2.BoundText)
    End If
     If BrnchesReg <> "-1" Then
        StrSQL = StrSQL & " AND Transactions.BranchId in( " & BrnchesReg & " )"
       End If
      If BrnchAct <> "-1" Then
        StrSQL = StrSQL & " AND Transactions.BranchId in( " & BrnchAct & " )"
       End If
    If Me.DcbBranch.Text <> "" And val(DcbBranch.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   Transactions.BranchId = " & val(Me.DcbBranch.BoundText)
    End If
    
    
    If Not IsNull(Me.DtpDateFrom2.value) Then
        StrSQL = StrSQL & " AND Transactions.Transaction_Date >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
    End If
    If Not IsNull(Me.DtpDateTo2.value) Then
        StrSQL = StrSQL & " AND Transactions.Transaction_Date<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
    End If
    
      If Not IsNull(Me.XPDtbTransTimeFrom.value) Then
       
                   StrSQL = StrSQL & " AND CAST(last_changed as time) >='" & FormatDateTime(Me.XPDtbTransTimeFrom.value, vbShortTime) & "'"
      End If
       If Not IsNull(Me.XPDtbTransTimeFrom.value) Then
                   StrSQL = StrSQL & " AND CAST(last_changed as time)<='" & FormatDateTime(Me.XPDtbTransTimeTo.value, vbShortTime) & "'"
      End If
      
If optPos(0).value = True Then
StrSQL = StrSQL & " AND   Transactions.POSBillType =1 "

ElseIf optPos(1).value = True Then
StrSQL = StrSQL & " AND    isnull(Transactions.POSBillType,0) =0 "
 
End If
  '  StrSQL = StrSQL & " ORDER BY last_changed,CAST(last_changed AS TIME)"
    GoTo xl:
        End If
        
    StrSQL = "select * from (( SELECT     POSBillType,dbo.Transactions.CashCustomerName,dbo.Transactions.PaymentType, dbo.TblTransactionPayments.id, dbo.TblTransactionPayments.Transaction_ID, dbo.TblTransactionPayments.PaymentID, ISNULL(dbo.TblTransactionPayments.[value], "
    StrSQL = StrSQL & "                   dbo.Transactions.Transaction_NetValue) AS Value, dbo.TblTransactionPayments.CardNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue,"
    StrSQL = StrSQL & "                    dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1,"
    StrSQL = StrSQL & "                   dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
    StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
    StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
    StrSQL = StrSQL & "                   dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Type, dbo.Transactions.StoreID,"
    StrSQL = StrSQL & "                   dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name,"
    StrSQL = StrSQL & "                   dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL & "        FROM         dbo.TblEmployee INNER JOIN"
    StrSQL = StrSQL & "                   dbo.Transactions ON dbo.TblEmployee.Emp_ID = dbo.Transactions.Emp_ID INNER JOIN"
    StrSQL = StrSQL & "                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "                   dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblTransactionPayments ON dbo.Transactions.Transaction_ID = dbo.TblTransactionPayments.Transaction_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID)"
    'where  "
    'StrSQL = StrSQL & "                  not(id is null)  and "
  '  StrSQL = StrSQL & "                  value>0)"
    StrSQL = StrSQL & " Union (SELECT   POSBillType,dbo.Transactions.CashCustomerName,dbo.Transactions.PaymentType,  dbo.TblSalesPayment.ID AS id, dbo.TblSalesPayment.TransID AS Transaction_ID, dbo.TblSalesPayment.PaymentID, ISNULL(dbo.TblSalesPayment.[Value], "
    StrSQL = StrSQL & "                  dbo.Transactions.Transaction_NetValue) AS value, dbo.TblSalesPayment.CardNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue,"
    StrSQL = StrSQL & "                  dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1,"
    StrSQL = StrSQL & "                  dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
    StrSQL = StrSQL & "                  dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
    StrSQL = StrSQL & "                  dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
    StrSQL = StrSQL & "                  dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Type, dbo.Transactions.StoreID,"
    StrSQL = StrSQL & "                  dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name,"
    StrSQL = StrSQL & "                  dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL & "         FROM         dbo.TblSalesPayment RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id ON"
    StrSQL = StrSQL & "                  dbo.TblSalesPayment.TransID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                  dbo.TblPaymentType ON dbo.TblSalesPayment.PaymentID = dbo.TblPaymentType.PaymentID where dbo.TblSalesPayment.PaymentID =0  )) as x where (Transaction_Type = 9 OR Transaction_Type = 21)"
    StrSQL = StrSQL & "                  dbo.TblPaymentType ON dbo.TblSalesPayment.PaymentID = dbo.TblPaymentType.PaymentID where    value>0  )"
   
StrSQL = StrSQL & "          Union ( "
StrSQL = StrSQL & "   SELECT     dbo.Transactions.POSBillType, dbo.Transactions.CashCustomerName, dbo.Transactions.PaymentType, dbo.TblTransactionPayments.id, "
StrSQL = StrSQL & "                        dbo.TblTransactionPayments.Transaction_ID, dbo.TblTransactionPayments.PaymentID, ISNULL(dbo.TblTransactionPayments.[value],"
StrSQL = StrSQL & "                        dbo.Transactions.Transaction_NetValue) AS Value, dbo.TblTransactionPayments.CardNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue,"
StrSQL = StrSQL & "                         dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1,"
StrSQL = StrSQL & "                        dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
StrSQL = StrSQL & "                        dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
StrSQL = StrSQL & "                        dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
StrSQL = StrSQL & "                        dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Type, dbo.Transactions.StoreID,"
StrSQL = StrSQL & "                        dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name,"
StrSQL = StrSQL & "                        dbo.TblBranchesData.branch_namee"
StrSQL = StrSQL & "  FROM         dbo.TblEmployee INNER JOIN"
StrSQL = StrSQL & "                        dbo.Transactions ON dbo.TblEmployee.Emp_ID = dbo.Transactions.Emp_ID INNER JOIN"
StrSQL = StrSQL & "                        dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblTransactionPayments ON dbo.Transactions.Transaction_ID = dbo.TblTransactionPayments.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
StrSQL = StrSQL & "   where (dbo.transactions.POSBillType Is Null) And (dbo.transactions.Transaction_Type = 9) and  isnull(transactions.PaymentType,0) <> 1"
StrSQL = StrSQL & "  )"
StrSQL = StrSQL & "  ) as x where (Transaction_Type = 9 OR Transaction_Type = 21) and  isnull(t.PaymentType,0) <> 1"
' StrSQL = StrSQL & " AND  not( Transaction_ID  is null)"
  
    If Me.DCboStoreName2.Text <> "" And val(DCboStoreName2.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   StoreID = " & val(Me.DCboStoreName2.BoundText)
    End If
     If BrnchesReg <> "-1" Then
        StrSQL = StrSQL & " AND BranchId in( " & BrnchesReg & " )"
       End If
      If BrnchAct <> "-1" Then
        StrSQL = StrSQL & " AND BranchId in( " & BrnchAct & " )"
       End If
    If Me.DcbBranch.Text <> "" And val(DcbBranch.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   BranchId = " & val(Me.DcbBranch.BoundText)
    End If
    If Me.CboPayMentType.Text <> "" And val(CboPayMentType.ListIndex) <> -1 Then
        StrSQL = StrSQL & " AND    PaymentID= " & val(Me.CboPayMentType.ListIndex)
    End If
    If Me.DcbEmp.Text <> "" And val(DcbEmp.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND    Emp_ID= " & val(Me.DcbEmp.BoundText)
    End If


    If Not IsNull(Me.DtpDateFrom2.value) Then
        StrSQL = StrSQL & " AND Transaction_Date >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
    End If
    If Not IsNull(Me.DtpDateTo2.value) Then
        StrSQL = StrSQL & " AND Transaction_Date<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
    End If
    
If optPos(0).value = True Then
StrSQL = StrSQL & " AND   POSBillType =1 "

ElseIf optPos(1).value = True Then
StrSQL = StrSQL & " AND    isnull(POSBillType,0) =0 "
 
End If
    Set rs = New ADODB.Recordset
    
    If optNetWork(6).value = True Then
 StrSQL = StrSQL & " AND         PaymentType<>1"
StrSQL = StrSQL & " AND    PaymentID = 0"
ElseIf optNetWork(7).value = True Then
 StrSQL = StrSQL & " AND         PaymentType<>1"
StrSQL = StrSQL & " AND    PaymentID in(2,5,7)"

 
ElseIf optNetWork(8).value = True Then
 StrSQL = StrSQL & " AND         PaymentType<>1"
StrSQL = StrSQL & " AND    PaymentID in(4,6,8)"

ElseIf optNetWork(9).value = True Then
 StrSQL = StrSQL & " AND         PaymentType=1"
  
    End If
    
'   Dim StrSQL As String
StrSQL = "WITH AllPayments AS ("

'------ TblTransactionPayments ------
StrSQL = StrSQL & " SELECT "
StrSQL = StrSQL & "     dbo.Transactions.POSBillType, dbo.Transactions.CashCustomerName, dbo.Transactions.PaymentType, "
StrSQL = StrSQL & "     dbo.TblTransactionPayments.id, dbo.TblTransactionPayments.Transaction_ID, dbo.TblTransactionPayments.PaymentID, "
StrSQL = StrSQL & "     ISNULL(dbo.TblTransactionPayments.[value], dbo.Transactions.Transaction_NetValue) AS Value, "
StrSQL = StrSQL & "     dbo.TblTransactionPayments.CardNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue, "
StrSQL = StrSQL & "     dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, "
StrSQL = StrSQL & "     dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, "
StrSQL = StrSQL & "     dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, "
StrSQL = StrSQL & "     dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, "
StrSQL = StrSQL & "     dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Type, dbo.Transactions.StoreID, "
StrSQL = StrSQL & "     dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "     dbo.TblBranchesData.branch_namee "
StrSQL = StrSQL & " FROM dbo.TblEmployee "
StrSQL = StrSQL & " INNER JOIN dbo.Transactions ON dbo.TblEmployee.Emp_ID = dbo.Transactions.Emp_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type "
StrSQL = StrSQL & " INNER JOIN dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID "
StrSQL = StrSQL & " INNER JOIN dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblTransactionPayments ON dbo.Transactions.Transaction_ID = dbo.TblTransactionPayments.Transaction_ID "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID "

StrSQL = StrSQL & " WHERE (dbo.Transactions.Transaction_Type = 9 OR dbo.Transactions.Transaction_Type = 21) "

'------ TblSalesPayment »œÊ‰  ﬂ—«— «·⁄„·Ì«  ------
StrSQL = StrSQL & " UNION "
StrSQL = StrSQL & " SELECT "
StrSQL = StrSQL & "     dbo.Transactions.POSBillType, dbo.Transactions.CashCustomerName, dbo.Transactions.PaymentType, "
StrSQL = StrSQL & "     dbo.TblSalesPayment.ID AS id, dbo.TblSalesPayment.TransID AS Transaction_ID, dbo.TblSalesPayment.PaymentID, "
StrSQL = StrSQL & "     ISNULL(dbo.TblSalesPayment.[Value], dbo.Transactions.Transaction_NetValue) AS Value, "
StrSQL = StrSQL & "     dbo.TblSalesPayment.CardNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue, "
StrSQL = StrSQL & "     dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, "
StrSQL = StrSQL & "     dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, "
StrSQL = StrSQL & "     dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, "
StrSQL = StrSQL & "     dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, "
StrSQL = StrSQL & "     dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Type, dbo.Transactions.StoreID, "
StrSQL = StrSQL & "     dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "     dbo.TblBranchesData.branch_namee "
StrSQL = StrSQL & " FROM dbo.TblSalesPayment "
StrSQL = StrSQL & " INNER JOIN dbo.Transactions ON dbo.TblSalesPayment.TransID = dbo.Transactions.Transaction_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type "
StrSQL = StrSQL & " INNER JOIN dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID "
StrSQL = StrSQL & " INNER JOIN dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id "
StrSQL = StrSQL & " LEFT OUTER JOIN dbo.TblPaymentType ON dbo.TblSalesPayment.PaymentID = dbo.TblPaymentType.PaymentID "
StrSQL = StrSQL & " WHERE (dbo.Transactions.Transaction_Type = 9 OR dbo.Transactions.Transaction_Type = 21) "
StrSQL = StrSQL & " AND NOT EXISTS (SELECT 1 FROM dbo.TblTransactionPayments tp WHERE tp.Transaction_ID = dbo.TblSalesPayment.TransID) "

'------ ⁄„·Ì«  «·‰ﬁœÌ ›ﬁÿ (Transaction ﬂ«‘ »œÊ‰ √Ì ”ÿ— œ›⁄) ------
StrSQL = StrSQL & " UNION "
StrSQL = StrSQL & " SELECT "
StrSQL = StrSQL & "     dbo.Transactions.POSBillType, dbo.Transactions.CashCustomerName, dbo.Transactions.PaymentType, "
StrSQL = StrSQL & "     NULL AS id, dbo.Transactions.Transaction_ID, 0 AS PaymentID, "
StrSQL = StrSQL & "     dbo.Transactions.Transaction_NetValue AS Value, "
StrSQL = StrSQL & "     NULL AS CardNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue, "
StrSQL = StrSQL & "     dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, "
StrSQL = StrSQL & "     N'‰ﬁœÌ' AS PaymentName, N'‰ﬁœÌ' AS PaymentNamee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, "
StrSQL = StrSQL & "     dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, "
StrSQL = StrSQL & "     dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, "
StrSQL = StrSQL & "     dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Type, dbo.Transactions.StoreID, "
StrSQL = StrSQL & "     dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "     dbo.TblBranchesData.branch_namee "
StrSQL = StrSQL & " FROM dbo.Transactions "
StrSQL = StrSQL & " INNER JOIN dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type "
StrSQL = StrSQL & " INNER JOIN dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID "
StrSQL = StrSQL & " INNER JOIN dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id "
StrSQL = StrSQL & " WHERE (dbo.Transactions.Transaction_Type = 9 OR dbo.Transactions.Transaction_Type = 21)  and  isnull(transactions.PaymentType,0) <> 1"
StrSQL = StrSQL & " AND NOT EXISTS (SELECT 1 FROM dbo.TblTransactionPayments tp WHERE tp.Transaction_ID = dbo.Transactions.Transaction_ID) "

StrSQL = StrSQL & " ) SELECT * FROM AllPayments WHERE 1=1 "


'--- √ﬂ„· «·›·« — ﬂ„« ÂÌ ---
If Me.DCboStoreName2.Text <> "" And val(DCboStoreName2.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   StoreID = " & val(Me.DCboStoreName2.BoundText)
End If
If BrnchesReg <> "-1" Then
    StrSQL = StrSQL & " AND BranchId in( " & BrnchesReg & " )"
End If
If BrnchAct <> "-1" Then
    StrSQL = StrSQL & " AND BranchId in( " & BrnchAct & " )"
End If
If Me.DcbBranch.Text <> "" And val(DcbBranch.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   BranchId = " & val(Me.DcbBranch.BoundText)
End If
If Me.CboPayMentType.Text <> "" And val(CboPayMentType.ListIndex) <> -1 Then
    StrSQL = StrSQL & " AND    PaymentID= " & val(Me.CboPayMentType.ListIndex)
End If
If Me.DcbEmp.Text <> "" And val(Me.DcbEmp.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND    Emp_ID= " & val(Me.DcbEmp.BoundText)
End If

If Not IsNull(Me.DtpDateFrom2.value) Then
    StrSQL = StrSQL & " AND Transaction_Date >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
End If
If Not IsNull(Me.DtpDateTo2.value) Then
    StrSQL = StrSQL & " AND Transaction_Date<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
End If

If optPos(0).value = True Then
    StrSQL = StrSQL & " AND   POSBillType =1 "
ElseIf optPos(1).value = True Then
    StrSQL = StrSQL & " AND    isnull(POSBillType,0) =0 "
End If

If optNetWork(6).value = True Then
    StrSQL = StrSQL & " AND         PaymentType<>1"
    StrSQL = StrSQL & " AND    PaymentID = 0"
ElseIf optNetWork(7).value = True Then
    StrSQL = StrSQL & " AND         PaymentType<>1"
    StrSQL = StrSQL & " AND    PaymentID in(2,5,7)"
ElseIf optNetWork(8).value = True Then
    StrSQL = StrSQL & " AND         PaymentType<>1"
    StrSQL = StrSQL & " AND    PaymentID in(4,6,8)"
ElseIf optNetWork(9).value = True Then
    StrSQL = StrSQL & " AND         PaymentType=1"
End If


StrSQL = "WITH AllPayments AS ("

'-- 1) «·⁄„·Ì«  «·„”Ã·… ›Ì TblTransactionPayments (ﬂ· «·√‰Ê«⁄) »œÊ‰  ﬂ—«— --
StrSQL = StrSQL & " SELECT DISTINCT "
StrSQL = StrSQL & "     tr.POSBillType, tr.CashCustomerName, tr.PaymentType, "
StrSQL = StrSQL & "     tp.id, tp.Transaction_ID, ISNULL(tp.PaymentID, 0) AS PaymentID, "
StrSQL = StrSQL & "     ISNULL(tp.[value], tr.Transaction_NetValue) AS Value, "
StrSQL = StrSQL & "     tp.CardNo, tr.Transaction_Date, tr.Transaction_NetValue, "
StrSQL = StrSQL & "     tr.Transaction_HijriDate, tr.Transaction_Serial, tr.NoteSerial, tr.NoteSerial1, "
StrSQL = StrSQL & "     ISNULL(pt.PaymentName, N'‰ﬁœÌ') AS PaymentName, pt.PaymentNamee, tr.Emp_ID, e.Emp_Name, e.Emp_Name1, "
StrSQL = StrSQL & "     e.Emp_Name2, e.Emp_Name3, e.Emp_Name4, e.Fullcode, e.Emp_Namee4, "
StrSQL = StrSQL & "     e.Emp_Namee3, e.Emp_Namee2, e.Emp_Namee1, e.Emp_Namee, "
StrSQL = StrSQL & "     tt.TransactionTypeName, tt.TransactionEnglishName, tr.Transaction_Type, tr.StoreID, "
StrSQL = StrSQL & "     s.StoreName, s.StoreNamee, s.Code, tr.BranchId, b.branch_name, "
StrSQL = StrSQL & "     b.branch_namee "
StrSQL = StrSQL & " FROM dbo.TblTransactionPayments tp "
StrSQL = StrSQL & " INNER JOIN dbo.Transactions tr ON tp.Transaction_ID = tr.Transaction_ID "
StrSQL = StrSQL & " LEFT JOIN dbo.TblPaymentType pt ON tp.PaymentID = pt.PaymentID "
StrSQL = StrSQL & " INNER JOIN dbo.TblEmployee e ON tr.Emp_ID = e.Emp_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TransactionTypes tt ON tr.Transaction_Type = tt.Transaction_Type "
StrSQL = StrSQL & " INNER JOIN dbo.TblStore s ON tr.StoreID = s.StoreID "
StrSQL = StrSQL & " INNER JOIN dbo.TblBranchesData b ON tr.BranchId = b.branch_id "
StrSQL = StrSQL & " WHERE (tr.Transaction_Type = 9 OR tr.Transaction_Type = 21) "

'-- 2) «·⁄„·Ì«  «··Ì „·Â«‘ √Ì ”ÿ— œ›⁄ Œ«·’ (‰ﬁœÌ ›ﬁÿ) --
StrSQL = StrSQL & " UNION ALL "
StrSQL = StrSQL & " SELECT "
StrSQL = StrSQL & "     tr.POSBillType, tr.CashCustomerName, tr.PaymentType, "
StrSQL = StrSQL & "     NULL AS id, tr.Transaction_ID, 0 AS PaymentID, "
StrSQL = StrSQL & "     tr.Transaction_NetValue AS Value, "
StrSQL = StrSQL & "     NULL AS CardNo, tr.Transaction_Date, tr.Transaction_NetValue, "
StrSQL = StrSQL & "     tr.Transaction_HijriDate, tr.Transaction_Serial, tr.NoteSerial, tr.NoteSerial1, "
StrSQL = StrSQL & "     N'‰ﬁœÌ' AS PaymentName, N'‰ﬁœÌ' AS PaymentNamee, tr.Emp_ID, e.Emp_Name, e.Emp_Name1, "
StrSQL = StrSQL & "     e.Emp_Name2, e.Emp_Name3, e.Emp_Name4, e.Fullcode, e.Emp_Namee4, "
StrSQL = StrSQL & "     e.Emp_Namee3, e.Emp_Namee2, e.Emp_Namee1, e.Emp_Namee, "
StrSQL = StrSQL & "     tt.TransactionTypeName, tt.TransactionEnglishName, tr.Transaction_Type, tr.StoreID, "
StrSQL = StrSQL & "     s.StoreName, s.StoreNamee, s.Code, tr.BranchId, b.branch_name, "
StrSQL = StrSQL & "     b.branch_namee "
StrSQL = StrSQL & " FROM dbo.Transactions tr "
StrSQL = StrSQL & " INNER JOIN dbo.TblEmployee e ON tr.Emp_ID = e.Emp_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TransactionTypes tt ON tr.Transaction_Type = tt.Transaction_Type "
StrSQL = StrSQL & " INNER JOIN dbo.TblStore s ON tr.StoreID = s.StoreID "
StrSQL = StrSQL & " INNER JOIN dbo.TblBranchesData b ON tr.BranchId = b.branch_id "
StrSQL = StrSQL & " WHERE (tr.Transaction_Type = 9 OR tr.Transaction_Type = 21) "
If optNetWork(9).value = True Then
    StrSQL = StrSQL & " and  isnull(tr.PaymentType,0) = 1"
Else
    StrSQL = StrSQL & " and  isnull(tr.PaymentType,0) <> 1"
End If
StrSQL = StrSQL & " AND NOT EXISTS (SELECT 1 FROM dbo.TblTransactionPayments tp WHERE tp.Transaction_ID = tr.Transaction_ID) "

StrSQL = StrSQL & " ) SELECT * FROM AllPayments WHERE 1=1 "
' --- Â‰« «·›·« — “Ì „« » Õ» »⁄œ WHERE 1=1 ---

If Me.DCboStoreName2.Text <> "" And val(Me.DCboStoreName2.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   StoreID = " & val(Me.DCboStoreName2.BoundText)
End If
If BrnchesReg <> "-1" Then
    StrSQL = StrSQL & " AND BranchId in( " & BrnchesReg & " )"
End If
If BrnchAct <> "-1" Then
    StrSQL = StrSQL & " AND BranchId in( " & BrnchAct & " )"
End If
If Me.DcbBranch.Text <> "" And val(Me.DcbBranch.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   BranchId = " & val(Me.DcbBranch.BoundText)
End If
If Me.CboPayMentType.Text <> "" And val(Me.CboPayMentType.ListIndex) <> -1 Then
    StrSQL = StrSQL & " AND    PaymentID= " & val(Me.CboPayMentType.ListIndex)
End If
If Me.DcbEmp.Text <> "" And val(Me.DcbEmp.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND    Emp_ID= " & val(Me.DcbEmp.BoundText)
End If

If Not IsNull(Me.DtpDateFrom2.value) Then
    StrSQL = StrSQL & " AND Transaction_Date >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
End If
If Not IsNull(Me.DtpDateTo2.value) Then
    StrSQL = StrSQL & " AND Transaction_Date<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
End If

If optPos(0).value = True Then
    StrSQL = StrSQL & " AND   POSBillType =1 "
ElseIf optPos(1).value = True Then
    StrSQL = StrSQL & " AND    isnull(POSBillType,0) =0 "
End If

If optNetWork(6).value = True Then
    StrSQL = StrSQL & " AND         PaymentType<>1"
    StrSQL = StrSQL & " AND    PaymentID = 0"
ElseIf optNetWork(7).value = True Then
    StrSQL = StrSQL & " AND         PaymentType<>1"
    StrSQL = StrSQL & " AND    PaymentID in(2,5,7)"
ElseIf optNetWork(8).value = True Then
    StrSQL = StrSQL & " AND         PaymentType<>1"
    StrSQL = StrSQL & " AND    PaymentID in(4,6,8)"
ElseIf optNetWork(9).value = True Then
    StrSQL = StrSQL & " AND         PaymentType=1"
End If



StrSQL = "WITH AllPayments AS ("

'-- 1) «·⁄„·Ì«  «·„”Ã·… ›Ì TblTransactionPayments (ﬂ· «·√‰Ê«⁄ »œÊ‰  ﬂ—«—) --
StrSQL = StrSQL & " SELECT DISTINCT "
StrSQL = StrSQL & "     tr.POSBillType, tr.CashCustomerName, tr.PaymentType, "
StrSQL = StrSQL & "     tp.id, tp.Transaction_ID, ISNULL(tp.PaymentID, 0) AS PaymentID, "
StrSQL = StrSQL & "     ISNULL(tp.[value], tr.Transaction_NetValue) AS Value, "
StrSQL = StrSQL & "     tp.CardNo, tr.Transaction_Date, tr.Transaction_NetValue, "
StrSQL = StrSQL & "     tr.Transaction_HijriDate, tr.Transaction_Serial, tr.NoteSerial, tr.NoteSerial1, "
StrSQL = StrSQL & "     ISNULL(pt.PaymentName, N'‰ﬁœÌ') AS PaymentName, pt.PaymentNamee, tr.Emp_ID, e.Emp_Name, e.Emp_Name1, "
StrSQL = StrSQL & "     e.Emp_Name2, e.Emp_Name3, e.Emp_Name4, e.Fullcode, e.Emp_Namee4, "
StrSQL = StrSQL & "     e.Emp_Namee3, e.Emp_Namee2, e.Emp_Namee1, e.Emp_Namee, "
StrSQL = StrSQL & "     tt.TransactionTypeName, tt.TransactionEnglishName, tr.Transaction_Type, tr.StoreID, "
StrSQL = StrSQL & "     s.StoreName, s.StoreNamee, s.Code, tr.BranchId, b.branch_name, "
StrSQL = StrSQL & "     b.branch_namee "
StrSQL = StrSQL & " FROM dbo.TblTransactionPayments tp "
StrSQL = StrSQL & " INNER JOIN dbo.Transactions tr ON tp.Transaction_ID = tr.Transaction_ID "
StrSQL = StrSQL & " LEFT JOIN dbo.TblPaymentType pt ON tp.PaymentID = pt.PaymentID "
StrSQL = StrSQL & " INNER JOIN dbo.TblEmployee e ON tr.Emp_ID = e.Emp_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TransactionTypes tt ON tr.Transaction_Type = tt.Transaction_Type "
StrSQL = StrSQL & " INNER JOIN dbo.TblStore s ON tr.StoreID = s.StoreID "
StrSQL = StrSQL & " INNER JOIN dbo.TblBranchesData b ON tr.BranchId = b.branch_id "
StrSQL = StrSQL & " WHERE (tr.Transaction_Type = 9 OR tr.Transaction_Type = 21) "

'-- 2) «·⁄„·Ì«  «··Ì „·Â«‘ √Ì ”ÿ— œ›⁄ Œ«’ (‰ﬁœÌ ›ﬁÿ) --
StrSQL = StrSQL & " UNION ALL "
StrSQL = StrSQL & " SELECT "
StrSQL = StrSQL & "     tr.POSBillType, tr.CashCustomerName, tr.PaymentType, "
StrSQL = StrSQL & "     NULL AS id, tr.Transaction_ID, 0 AS PaymentID, "
StrSQL = StrSQL & "     tr.Transaction_NetValue AS Value, "
StrSQL = StrSQL & "     NULL AS CardNo, tr.Transaction_Date, tr.Transaction_NetValue, "
StrSQL = StrSQL & "     tr.Transaction_HijriDate, tr.Transaction_Serial, tr.NoteSerial, tr.NoteSerial1, "
StrSQL = StrSQL & "     N'‰ﬁœÌ' AS PaymentName, N'‰ﬁœÌ' AS PaymentNamee, tr.Emp_ID, e.Emp_Name, e.Emp_Name1, "
StrSQL = StrSQL & "     e.Emp_Name2, e.Emp_Name3, e.Emp_Name4, e.Fullcode, e.Emp_Namee4, "
StrSQL = StrSQL & "     e.Emp_Namee3, e.Emp_Namee2, e.Emp_Namee1, e.Emp_Namee, "
StrSQL = StrSQL & "     tt.TransactionTypeName, tt.TransactionEnglishName, tr.Transaction_Type, tr.StoreID, "
StrSQL = StrSQL & "     s.StoreName, s.StoreNamee, s.Code, tr.BranchId, b.branch_name, "
StrSQL = StrSQL & "     b.branch_namee "
StrSQL = StrSQL & " FROM dbo.Transactions tr "
StrSQL = StrSQL & " INNER JOIN dbo.TblEmployee e ON tr.Emp_ID = e.Emp_ID "
StrSQL = StrSQL & " INNER JOIN dbo.TransactionTypes tt ON tr.Transaction_Type = tt.Transaction_Type "
StrSQL = StrSQL & " INNER JOIN dbo.TblStore s ON tr.StoreID = s.StoreID "
StrSQL = StrSQL & " INNER JOIN dbo.TblBranchesData b ON tr.BranchId = b.branch_id "
StrSQL = StrSQL & " WHERE (tr.Transaction_Type = 9 OR tr.Transaction_Type = 21) "
If optNetWork(9).value = True Then
    StrSQL = StrSQL & " and  isnull(tr.PaymentType,0) = 1"
Else
    StrSQL = StrSQL & " and  isnull(tr.PaymentType,0) <> 1"
End If
StrSQL = StrSQL & " AND NOT EXISTS (SELECT 1 FROM dbo.TblTransactionPayments tp WHERE tp.Transaction_ID = tr.Transaction_ID) "

'--- SELECT «·‰Â«∆Ì „⁄ ≈÷«›… creditTotal ---
StrSQL = StrSQL & " ) SELECT *, CASE WHEN PaymentType = 1 THEN Value ELSE 0 END AS creditTotal FROM AllPayments WHERE 1=1 "

If Me.DCboStoreName2.Text <> "" And val(Me.DCboStoreName2.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   StoreID = " & val(Me.DCboStoreName2.BoundText)
End If
If BrnchesReg <> "-1" Then
    StrSQL = StrSQL & " AND BranchId in( " & BrnchesReg & " )"
End If
If BrnchAct <> "-1" Then
    StrSQL = StrSQL & " AND BranchId in( " & BrnchAct & " )"
End If
If Me.DcbBranch.Text <> "" And val(Me.DcbBranch.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   BranchId = " & val(Me.DcbBranch.BoundText)
End If
If Me.CboPayMentType.Text <> "" And val(Me.CboPayMentType.ListIndex) <> -1 Then
    StrSQL = StrSQL & " AND    PaymentID= " & val(Me.CboPayMentType.ListIndex)
End If
If Me.DcbEmp.Text <> "" And val(Me.DcbEmp.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND    Emp_ID= " & val(Me.DcbEmp.BoundText)
End If

If Not IsNull(Me.DtpDateFrom2.value) Then
    StrSQL = StrSQL & " AND Transaction_Date >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
End If
If Not IsNull(Me.DtpDateTo2.value) Then
    StrSQL = StrSQL & " AND Transaction_Date<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
End If

If optPos(0).value = True Then
    StrSQL = StrSQL & " AND   POSBillType =1 "
ElseIf optPos(1).value = True Then
    StrSQL = StrSQL & " AND    isnull(POSBillType,0) =0 "
End If

If optNetWork(6).value = True Then
    StrSQL = StrSQL & " AND         PaymentType<>1"
    StrSQL = StrSQL & " AND    PaymentID = 0"
ElseIf optNetWork(7).value = True Then
    StrSQL = StrSQL & " AND         PaymentType<>1"
    StrSQL = StrSQL & " AND    PaymentID in(2,5,7)"
ElseIf optNetWork(8).value = True Then
    StrSQL = StrSQL & " AND         PaymentType<>1"
    StrSQL = StrSQL & " AND    PaymentID in(4,6,8)"
ElseIf optNetWork(9).value = True Then
    StrSQL = StrSQL & " AND         PaymentType=1"
End If

' «·¬‰ «·‹ StrSQL ÌÕ ÊÌ creditTotal Ê ﬁœ—  Ã„⁄Â √Ê  ” Œœ„Â ›Ì «· ﬁ«—Ì— „»«‘—…


Set rs = New ADODB.Recordset
 
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
 rs.MoveFirst
xl:
 If optNetWork(0).value = True Then
 print_report StrSQL, 1
 ElseIf optNetWork(1).value = True Then
 print_report StrSQL, 2
  ElseIf optNetWork(2).value = True Then
 print_report StrSQL, 3
 
   ElseIf optNetWork(5).value = True Then
 print_report StrSQL, 6
  
   ElseIf optNetWork(11).value = True Then
 print_report StrSQL, 11
 
    ElseIf optNetWork(6).value = True Then
 print_report StrSQL, 7
 
 
    ElseIf optNetWork(7).value = True Then
 print_report StrSQL, 8
 
 
    ElseIf optNetWork(8).value = True Then
 print_report StrSQL, 9
 
 
    ElseIf optNetWork(9).value = True Then
 print_report StrSQL, 10
  
  
 
   ElseIf optNetWork(3).value = True Then
   StrSQL = StrSQL & " ORDER BY Transaction_Date,CAST(last_changed AS TIME) ASC"
 print_report StrSQL, 4
   ElseIf optNetWork(4).value = True Then
 print_report StrSQL, 5
 
End If
    End If
End Sub
Function print_report2(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     Debug.Print NoteSerial
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    If Ind = 1 Then
    If Optrans(2).value = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWorkSales.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWorkSales.rpt"
         End If
   Else
          If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWork.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWork.rpt"
          End If
   End If
   ElseIf Ind = 2 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWork2.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWork2.rpt"
            
       End If
 End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
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

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
  
    End If

   If Ind = 2 Or Ind = 1 Then
  If Not IsNull(DTPicker1.value) And Not IsNull(DTPicker2.value) Then
   xReport.ParameterFields(8).AddCurrentValue DTPicker1.value
    xReport.ParameterFields(10).AddCurrentValue DTPicker2.value
    End If
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
Function print_report(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     Debug.Print NoteSerial
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    If Ind = 1 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkTotal.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkTotalE.rpt"
            
       End If
     ElseIf Ind = 2 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnaly.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalyE.rpt"
            
       End If
         
         
        ElseIf Ind = 50 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EmpCommission2.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EmpCommission2.rpt"
            
       End If
  
         
         ElseIf Ind = 3 Then
         
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2E.rpt"
            
       End If
       
       
             ElseIf Ind = 4 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "kpi1.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "kpi1.rpt"
            
       End If
       
        ElseIf Ind = 5 Then
         If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "kpi2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "kpi2.rpt"
            
        End If
       
       
                   ElseIf Ind = 6 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2days.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2days.rpt"
            
       End If
       
       
       
                   ElseIf Ind = 11 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysShort.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysShort.rpt"
            
       End If
       
       
       
'********************************************

                   ElseIf Ind = 7 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2dayscash.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2dayscash.rpt"
            
       End If
      
                   ElseIf Ind = 8 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysMada.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysMada.rpt"
            
       End If
       
                          ElseIf Ind = 9 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysvisa.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysvisa.rpt"
            
       End If
       
      
                         ElseIf Ind = 10 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysCredit.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysvisaCredit.rpt"
            
       End If
        
        
  '********************************************
       
       
    Else
    
   If reportid = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItems.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItems.rpt"
            
       End If
       
       If ChsERIAL.value = vbChecked Then
       StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItemsSerials.rpt"
       End If
       
      ElseIf reportid = 15 Then '«·ÿ·»Ì«  «·ÃœÌœ…
      
      
      
            If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "DetailsOrderNew.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "DetailsOrderNew.rpt"
            
       End If
       
       
      
Else


       If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItems1.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItems1.rpt"
            
       End If
       
       If ChsERIAL.value = vbChecked Then
       StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItems1Serials.rpt"
       End If
              
              
End If
End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

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

Dim i As Integer
For i = 0 To 14
 If opt(i).value = True Then
 StrReportTitle = opt(i).Caption
 
 If DBCboClientName.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··⁄„Ì· : " & DBCboClientName.Text
 End If
 
 
 If DCboStoreName.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··„Œ“‰ : " & DCboStoreName.Text
 End If
 
  If DCboGroup1.Text <> "" Then
StrReportTitle = StrReportTitle & CHR(13) & "  ··„Ã„Ê⁄Â : " & DCboGroup1.Text
 End If
 
 If ItemDetailedCode.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··ﬂÊœ : " & ItemDetailedCode.Text
 End If
 
  If ParrtNoCode.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··»«—ﬂÊœ : " & ParrtNoCode.Text
 End If
 
   If DcbColor.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··Ê‰ : " & DcbColor.Text
 End If
 
 
   If DcbSize.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··„ﬁ«” : " & DcbSize.Text
 End If
 
 
 
 
 
 End If
 
 
  
Next i


    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        If reportid = 15 Then
        xReport.ParameterFields(12).AddCurrentValue val(percent1.Text)
        xReport.ParameterFields(13).AddCurrentValue val(percent2.Text)
        
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
 '       StrReportTitle = ""
  
    End If

   If Ind = 0 Then
  If Not IsNull(DtpDateFrom.value) And Not IsNull(DtpDateTo.value) Then
   xReport.ParameterFields(8).AddCurrentValue DtpDateFrom.value
    xReport.ParameterFields(10).AddCurrentValue DtpDateTo.value
    End If
Else
  
 If Ind <> 50 Then
  If Not IsNull(DtpDateFrom2.value) And Not IsNull(DtpDateTo2.value) Then
   xReport.ParameterFields(8).AddCurrentValue DtpDateFrom2.value
    xReport.ParameterFields(10).AddCurrentValue DtpDateTo2.value
    End If
    End If
End If
  Dim total As String
  Dim totl As Double


    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , NoteSerial

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function




Private Sub UpdateZatcaStatus(ByVal mInvoiceId As String)
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim FilePath As String
    Dim lastrow As Long
    Dim i As Long
    Dim invoiceID As String
    Dim Col
    Dim zatcaStatusCol As Long
Dim found As Boolean
    '  ÕœÌœ „”«— „·› Excel
    FilePath = excelFileNameFullPath

    ' ≈‰‘«¡ ﬂ«∆‰ Excel ÃœÌœ »«” Œœ«„ Late Binding
    Set xlApp = CreateObject("Excel.Application")

    ' › Õ „·› Excel
    Set xlBook = xlApp.Workbooks.Open(FilePath)

    '  ÕœÌœ «·‘Ì  «·√Ê· »€÷ «·‰Ÿ— ⁄‰ «”„Â
    Set xlSheet = xlBook.Sheets(1)

    '  ÕœÌœ ¬Œ— ’› ÌÕ ÊÌ ⁄·Ï »Ì«‰« 
    lastrow = xlSheet.cells(xlSheet.rows.count, "A").End(-4162).Row ' -4162 ÂÊ xlUp ›Ì VB6

    '  ÕœÌœ ⁄„Êœ zatcaStatus
    zatcaStatusCol = xlSheet.cells(1, xlSheet.Columns.count).End(-4159).Column + 1 ' -4159 ÂÊ xlToLeft ›Ì VB6
   ' xlSheet.Cells(1, zatcaStatusCol).value = "zatcaStatus"
    found = False
    For Col = 1 To xlSheet.UsedRange.Columns.count
        If xlSheet.cells(1, Col).value = "zatcaStatus" Then
            zatcaStatusCol = Col
            found = True
            Exit For
        End If
    Next Col

    ' ≈–« ·„ Ì „ «·⁄ÀÊ— ⁄·Ï «·⁄„Êœ° √÷›Â
    If Not found Then
        zatcaStatusCol = xlSheet.cells(1, xlSheet.Columns.count).End(-4159).Column + 1 ' -4159 ÂÊ xlToLeft ›Ì VB6
        xlSheet.cells(1, zatcaStatusCol).value = "zatcaStatus"
    End If
    

    ' «·„—Ê— ⁄»— «·’›Ê› Ê ÕœÌÀ zatcaStatus »‰«¡ ⁄·Ï „⁄«ÌÌ— „⁄Ì‰…
    For i = 2 To lastrow ' Ì› —÷ √‰ «·’› «·√Ê· ÌÕ ÊÌ ⁄·Ï —ƒÊ” «·√⁄„œ…
        invoiceID = xlSheet.cells(i, 2).value ' «” »œ· 1 »—ﬁ„ «·⁄„Êœ «·–Ì ÌÕ ÊÌ ⁄·Ï invoiceID

        '  ÿ»Ìﬁ «·„⁄«ÌÌ— «·„ÿ·Ê»…
        If invoiceID = mInvoiceId Then ' ÷⁄ «·„⁄Ì«— «·Œ«’ »ﬂ Â‰«
            xlSheet.cells(i, zatcaStatusCol).value = 1
        Else
            xlSheet.cells(i, zatcaStatusCol).value = 0
        End If
    Next i

    ' Õ›Ÿ Ê≈€·«ﬁ „·› Excel
    xlBook.save
    xlBook.Close False

    ' ≈‰Â«¡ Excel
    xlApp.Quit

    '  ‰ŸÌ› «·ﬂ«∆‰« 
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

   ' MsgBox "Data updated successfully!"
End Sub


Private Sub UpdateZatcaStatus2(ByVal mInvoiceId As String)
'    Dim xlApp As Excel.Application
'    Dim xlBook As Excel.Workbook
'    Dim xlSheet As Excel.Worksheet
'    Dim FilePath As String
'    Dim lastrow As Long
'    Dim i As Long
'    Dim invoiceID As String
'    Dim zatcaStatusCol As Long
'
'    ' ????? ???? ??? Excel
'    FilePath = excelFileNameFullPath ' ?????? ?????? ????? ??? Excel ????? ??
'
'    ' ????? ???? Excel ????
'    Set xlApp = New Excel.Application
'
'    ' ??? ??? Excel
'    Set xlBook = xlApp.Workbooks.Open(FilePath)
'
'    ' ????? ????? ????? ??? ????? ?? ????
'    Set xlSheet = xlBook.Sheets(1)
'
'    ' ????? ??? ?? ????? ??? ??????
'    lastrow = xlSheet.cells(xlSheet.rows.count, "A").End(-4162).row ' -4162 ?? xlUp ?? VB6
'
'    ' ????? ???? zatcaStatus
'    zatcaStatusCol = xlSheet.cells(1, xlSheet.Columns.count).End(-4159).Column + 1 ' -4159 ?? xlToLeft ?? VB6
'    xlSheet.cells(1, zatcaStatusCol).value = "zatcaStatus"
'
'    ' ?????? ??? ?????? ?????? zatcaStatus ????? ??? ?????? ?????
'    For i = 2 To lastrow ' ????? ?? ???? ????? ????? ??? ???? ???????
'        invoiceID = xlSheet.cells(i, 1).value ' ?????? 1 ???? ?????? ???? ????? ??? invoiceID
'
'        ' ????? ???????? ????????
'        If invoiceID = mInvoiceId Then ' ?? ??????? ????? ?? ???
'            xlSheet.cells(i, zatcaStatusCol).value = 1
'        End If
'    Next i
'
'    ' ??? ?????? ??? Excel
'    xlBook.save
'    xlBook.Close False
'
'    ' ????? Excel
'    xlApp.Quit
'
'    ' ????? ????????
'    Set xlSheet = Nothing
'    Set xlBook = Nothing
'    Set xlApp = Nothing
'
'    MsgBox "Data updated successfully!"
End Sub

